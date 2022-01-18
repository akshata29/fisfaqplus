using Microsoft.ApplicationInsights;
using Microsoft.Azure.CognitiveServices.Knowledge.QnAMaker.Models;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Schema;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Teams.Apps.FAQPlusPlus.Bots;
using Microsoft.Teams.Apps.FAQPlusPlus.Cards;
using Microsoft.Teams.Apps.FAQPlusPlus.Common;
using Microsoft.Teams.Apps.FAQPlusPlus.Common.Models;
using Microsoft.Teams.Apps.FAQPlusPlus.Common.Models.Configuration;
using Microsoft.Teams.Apps.FAQPlusPlus.Common.Providers;
using Microsoft.Teams.Apps.FAQPlusPlus.Helpers;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Threading;
using System.Threading.Tasks;

namespace Microsoft.Teams.Apps.FAQPlusPlus.Dialogs
{
    public class PersonalChatMainDialog : ComponentDialog
    {
        private ChangeLanguageDialog _changeLangDialog;
        private readonly IQnaServiceProvider _qnaServiceProvider;
        private readonly string _appBaseUri;
        private readonly BotSettings _options;
        private readonly TelemetryClient _telemetryClient;
        private readonly IConfigurationDataProvider _configurationProvider;

        public PersonalChatMainDialog(ChangeLanguageDialog changeLangDialog,
            UserState userState,
            IQnaServiceProvider qnaServiceProvider,
            IOptionsMonitor<BotSettings> optionsAccessor,
            TelemetryClient telemetryClient,
            IConfigurationDataProvider configurationProvider)
            : base("root")
        {
            _telemetryClient = telemetryClient;
            _configurationProvider = configurationProvider;
            _changeLangDialog = changeLangDialog;
            _qnaServiceProvider = qnaServiceProvider;
            _options = optionsAccessor.CurrentValue;
            _appBaseUri = _options.AppBaseUri;

            AddDialog(changeLangDialog);
            AddDialog(new WaterfallDialog("personalRoot", new WaterfallStep[] { StartDialogAsync }));
            InitialDialogId = "personalRoot";
        }

        private async Task<DialogTurnResult> StartDialogAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            string text = stepContext.Context.Activity.Text?.ToLower()?.Trim() ?? string.Empty;

            switch (text)
            {
                case Constants.ChangeLanguage:
                    this._telemetryClient.TrackEvent("User wants to change language");
                    return await stepContext.BeginDialogAsync(nameof(ChangeLanguageDialog));
                    break;
                case Constants.AskAnExpert:
                    this._telemetryClient.TrackEvent("Sending user ask an expert card");
                    await stepContext.Context.SendActivityAsync(MessageFactory.Attachment(AskAnExpertCard.GetCard())).ConfigureAwait(false);
                    break;

                case Constants.ShareFeedback:
                    this._telemetryClient.TrackEvent("Sending user feedback card");
                    await stepContext.Context.SendActivityAsync(MessageFactory.Attachment(ShareFeedbackCard.GetCard())).ConfigureAwait(false);
                    break;

                case Constants.TakeATour:
                    this._telemetryClient.TrackEvent("Sending user tour card");
                    var userTourCards = TourCarousel.GetUserTourCards(this._appBaseUri);
                    await stepContext.Context.SendActivityAsync(MessageFactory.Carousel(userTourCards)).ConfigureAwait(false);
                    break;

                default:
                    this._telemetryClient.TrackEvent("Sending input to QnAMaker");
                    var message = stepContext.Context.Activity.AsMessageActivity();

                    await this.GetQuestionAnswerReplyAsync(stepContext.Context, message).ConfigureAwait(false);
                    break;
            }

            return await stepContext.EndDialogAsync(null, cancellationToken);
        }

        /// <summary>
        /// Get the reply to a question asked by end user.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="message">Text message.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task GetQuestionAnswerReplyAsync(
            ITurnContext turnContext,
            IMessageActivity message)
        {
            string text = message.Text?.ToLower()?.Trim() ?? string.Empty;

            try
            {
                var queryResult = new QnASearchResultList();

                ResponseCardPayload payload = new ResponseCardPayload();

                if (!string.IsNullOrEmpty(message.ReplyToId) && (message.Value != null))
                {
                    payload = ((JObject)message.Value).ToObject<ResponseCardPayload>();
                }

                queryResult = await _qnaServiceProvider.GenerateAnswerAsync(question: text, isTestKnowledgeBase: false, payload.PreviousQuestions?.First().Id.ToString(), payload.PreviousQuestions?.First().Questions.First()).ConfigureAwait(false);

                if (queryResult.Answers.First().Id != -1)
                {
                    var answerData = queryResult.Answers.First();
                    payload.QnaPairId = answerData.Id ?? -1;

                    AnswerModel answerModel = new AnswerModel();

                    if (Validators.IsValidJSON(answerData.Answer))
                    {
                        answerModel = JsonConvert.DeserializeObject<AnswerModel>(answerData.Answer);
                    }

                    if (!string.IsNullOrEmpty(answerModel?.Title) || !string.IsNullOrEmpty(answerModel?.Subtitle) || !string.IsNullOrEmpty(answerModel?.ImageUrl) || !string.IsNullOrEmpty(answerModel?.RedirectionUrl))
                    {
                        await turnContext.SendActivityAsync(MessageFactory.Attachment(MessagingExtensionQnaCard.GetEndUserRichCard(text, answerData, payload.QnaPairId))).ConfigureAwait(false);
                    }
                    else
                    {
                        await turnContext.SendActivityAsync(MessageFactory.Attachment(ResponseCard.GetCard(answerData, text, _appBaseUri, payload))).ConfigureAwait(false);
                    }

                    _telemetryClient.TrackEvent(
                        FaqPlusPlusBot.EVENT_ANSWERED_QUESTION_SINGLE,
                        new Dictionary<string, string>
                        {
                                { "QuestionId" ,payload.QnaPairId.ToString() },
                                { "QuestionAnswered", queryResult.Answers[0].Questions[0] },
                                { "QuestionAsked", text },
                                { "UserName" ,turnContext.Activity.From.Name},
                                { "UserAadId", turnContext.Activity.From?.AadObjectId ?? "" },
                                { "Product", _options.ProductName },
                        });

                }
                else
                {
                    await turnContext.SendActivityAsync(MessageFactory.Attachment(UnrecognizedInputCard.GetCard(text))).ConfigureAwait(false);
                }
            }
            catch (Exception ex)
            {
                // Check if knowledge base is empty and has not published yet when end user is asking a question to bot.
                if (((Azure.CognitiveServices.Knowledge.QnAMaker.Models.ErrorResponseException)ex).Response.StatusCode == HttpStatusCode.BadRequest)
                {
                    var knowledgeBaseId = await _configurationProvider.GetSavedEntityDetailAsync(Constants.KnowledgeBaseEntityId).ConfigureAwait(false);
                    var hasPublished = await _qnaServiceProvider.GetInitialPublishedStatusAsync(knowledgeBaseId).ConfigureAwait(false);

                    // Check if knowledge base has not published yet.
                    if (!hasPublished)
                    {
                        this._telemetryClient.TrackException(ex, new Dictionary<string, string> {
                            {
                                "message", "Error while fetching the qna pair: knowledge base may be empty or it has not published yet." },
                            });
                        await turnContext.SendActivityAsync(MessageFactory.Attachment(UnrecognizedInputCard.GetCard(text))).ConfigureAwait(false);
                        return;
                    }
                }

                // Throw the error at calling place, if there is any generic exception which is not caught.
                throw;
            }
        }

    }
}
