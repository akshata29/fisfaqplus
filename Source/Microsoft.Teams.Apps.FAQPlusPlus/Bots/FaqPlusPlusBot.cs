// <copyright file="FaqPlusPlusBot.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.FAQPlusPlus.Bots
{
    using System;
    using System.Collections.Concurrent;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Net;
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Azure.CognitiveServices.Knowledge.QnAMaker;
    using Microsoft.Azure.CognitiveServices.Knowledge.QnAMaker.Models;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Dialogs;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Connector;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Caching.Memory;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.FAQPlusPlus.Cards;
    using Microsoft.Teams.Apps.FAQPlusPlus.Common;
    using Microsoft.Teams.Apps.FAQPlusPlus.Common.Models;
    using Microsoft.Teams.Apps.FAQPlusPlus.Common.Models.Configuration;
    using Microsoft.Teams.Apps.FAQPlusPlus.Common.Providers;
    using Microsoft.Teams.Apps.FAQPlusPlus.Dialogs;
    using Microsoft.Teams.Apps.FAQPlusPlus.Helpers;
    using Microsoft.Teams.Apps.FAQPlusPlus.Properties;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;
    using ErrorResponseException = Microsoft.Azure.CognitiveServices.Knowledge.QnAMaker.Models.ErrorResponseException;

    /// <summary>
    /// Class that handles the teams activity of Faq Plus bot and messaging extension.
    /// </summary>
    public class FaqPlusPlusBot : TeamsActivityHandler
    {
        /// <summary>
        ///  Default access cache expiry in days to check if user using the app is a valid SME or not.
        /// </summary>
        private const int DefaultAccessCacheExpiryInDays = 5;

        /// <summary>
        /// Search text parameter name in the manifest file.
        /// </summary>
        private const string SearchTextParameterName = "searchText";

        /// <summary>
        /// Represents the task module height.
        /// </summary>
        private const int TaskModuleHeight = 450;

        /// <summary>
        /// Represents the task module width.
        /// </summary>
        private const int TaskModuleWidth = 500;

        /// <summary>
        /// Represents the conversation type as personal.
        /// </summary>
        private const string ConversationTypePersonal = "personal";

        /// <summary>
        ///  Represents the conversation type as channel.
        /// </summary>
        private const string ConversationTypeChannel = "channel";

        /// <summary>
        /// ChangeStatus - text that triggers change status action by SME.
        /// </summary>
        private const string ChangeStatus = "change status";

        /// <summary>
        /// File extension for CSV files
        /// </summary>
        private const string CSV_EXTENSION = ".csv";

        /// <summary>
        /// File extension for Excel files
        /// </summary>
        private const string XLSX_EXTENSION = ".xlsx";
        private const string XLSX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        private const string CSV_MIME_TYPE = "text/csv";

        // telemetry event names
        private const string EVENT_ANSWERED_QUESTION_BULK = "QuestionAnsweredBulk";
        private const string EVENT_TRANSLATED_QUESTION_BULK = "QuestionTranslatedBulk";
        private const string EVENT_TRANSLATED_ANSWERS_BULK = "QuestionTranslatedBulk";
        private const string EVENT_UPDATED_QUESTION = "QuestionUpdated";
        private const string EVENT_MESSAGE_RECEIVED = "MessageReceived";
        private const string EVENT_QUESTION_ADDED = "QuestionAdded";
        public const string EVENT_ANSWERED_QUESTION_SINGLE = "QuestionAnsweredSingle";
        public const string EVENT_LANGUAGE_PREFERENCE_CHANGED = "LanguagePreferenceChanged";

        /// <summary>
        /// Represents a set of key/value application configuration properties for FaqPlusPlus bot.
        /// </summary>
        private readonly BotSettings options;

        private readonly IConfigurationDataProvider configurationProvider;
        private readonly MicrosoftAppCredentials microsoftAppCredentials;
        private readonly ITicketsProvider ticketsProvider;
        private readonly IBatchFileProvider batchFileProvider;
        private readonly IActivityStorageProvider activityStorageProvider;
        private readonly ISearchService searchService;
        private readonly string appId;
        private readonly BotFrameworkAdapter botAdapter;
        private readonly IMemoryCache accessCache;
        private readonly int accessCacheExpiryInDays;
        private readonly string appBaseUri;
        private readonly IKnowledgeBaseSearchService knowledgeBaseSearchService;
        private readonly ILogger<FaqPlusPlusBot> logger;
        private readonly TelemetryClient telemetryClient;
        private readonly IQnaServiceProvider qnaServiceProvider;
        private readonly IHttpClientFactory clientFactory;
        private readonly IStatePropertyAccessor<string> languagePreference;
        private readonly UserState userState;
        private readonly ConversationState conversationState;
        private readonly PersonalChatMainDialog rootPersonalDialog;
        private readonly TranslatorService translatorService;
        private const string EnglishEnglish = "en";
        private const string EnglishSpanish = "es";
        private const string SpanishEnglish = "in";
        private const string SpanishSpanish = "it";
        private const string BatchResultsDirectory = "results/";

        /// <summary>
        /// Initializes a new instance of the <see cref="FaqPlusPlusBot"/> class.
        /// </summary>
        /// <param name="clientFactory">Client Factory</param>
        /// <param name="configurationProvider">Configuration Provider.</param>
        /// <param name="microsoftAppCredentials">Microsoft app credentials to use.</param>
        /// <param name="ticketsProvider">Tickets Provider.</param>
        /// <param name="activityStorageProvider">Activity storage provider.</param>
        /// <param name="qnaServiceProvider">Question and answer maker service provider.</param>
        /// <param name="searchService">SearchService dependency injection.</param>
        /// <param name="botAdapter">Bot adapter dependency injection.</param>
        /// <param name="memoryCache">IMemoryCache dependency injection.</param>
        /// <param name="knowledgeBaseSearchService">KnowledgeBaseSearchService dependency injection.</param>
        /// <param name="optionsAccessor">A set of key/value application configuration properties for FaqPlusPlus bot.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public FaqPlusPlusBot(
            IHttpClientFactory clientFactory,
            Common.Providers.IConfigurationDataProvider configurationProvider,
            MicrosoftAppCredentials microsoftAppCredentials,
            ITicketsProvider ticketsProvider,
            IBatchFileProvider batchFileStorageProvider,
            IQnaServiceProvider qnaServiceProvider,
            IActivityStorageProvider activityStorageProvider,
            ISearchService searchService,
            BotFrameworkAdapter botAdapter,
            IMemoryCache memoryCache,
            IKnowledgeBaseSearchService knowledgeBaseSearchService,
            IOptionsMonitor<BotSettings> optionsAccessor,
            ILogger<FaqPlusPlusBot> logger,
            UserState userState,
            PersonalChatMainDialog rootPersonalDialog,
            ConversationState conversationState,
            TelemetryClient telemetryClient,
            TranslatorService translatorService)
        {
            this.clientFactory = clientFactory;
            this.configurationProvider = configurationProvider;
            this.microsoftAppCredentials = microsoftAppCredentials;
            this.ticketsProvider = ticketsProvider;
            this.batchFileProvider = batchFileStorageProvider;
            this.options = optionsAccessor.CurrentValue;
            this.qnaServiceProvider = qnaServiceProvider;
            this.activityStorageProvider = activityStorageProvider;
            this.searchService = searchService;
            this.appId = this.options.MicrosoftAppId;
            this.botAdapter = botAdapter;
            this.accessCache = memoryCache;
            this.logger = logger;
            this.telemetryClient = telemetryClient;
            this.accessCacheExpiryInDays = this.options.AccessCacheExpiryInDays;

            if (this.accessCacheExpiryInDays <= 0)
            {
                this.accessCacheExpiryInDays = DefaultAccessCacheExpiryInDays;
                this.logger.LogInformation($"Configuration option is not present or out of range for AccessCacheExpiryInDays and the default value is set to: {this.accessCacheExpiryInDays}", SeverityLevel.Information);
            }

            this.appBaseUri = this.options.AppBaseUri;
            this.knowledgeBaseSearchService = knowledgeBaseSearchService;

            this.userState = userState ?? throw new NullReferenceException(nameof(userState));
            this.conversationState = conversationState ?? throw new NullReferenceException(nameof(conversationState));
            this.rootPersonalDialog = rootPersonalDialog;
            this.languagePreference = userState.CreateProperty<string>(TranslationMiddleware.PreferredLanguageSetting);
            this.translatorService = translatorService;
        }

        /// <summary>
        /// Handles an incoming activity.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        /// <remarks>
        /// Reference link: https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.activityhandler.onturnasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        public override async Task OnTurnAsync(
            ITurnContext turnContext,
            CancellationToken cancellationToken = default)
        {
            try
            {
                if (turnContext != null & !this.IsActivityFromExpectedTenant(turnContext))
                {
                    this.logger.LogWarning($"Unexpected tenant id {turnContext?.Activity.Conversation.TenantId}");
                    return;
                }

                // Get the current culture info to use in resource files
                string locale = turnContext?.Activity.Entities?.FirstOrDefault(entity => entity.Type == "clientInfo")?.Properties["locale"]?.ToString();

                if (!string.IsNullOrEmpty(locale))
                {
                    CultureInfo.CurrentCulture = CultureInfo.CurrentUICulture = CultureInfo.GetCultureInfo(locale);
                }

                await base.OnTurnAsync(turnContext, cancellationToken);


                await conversationState.SaveChangesAsync(turnContext, false, cancellationToken);
                await userState.SaveChangesAsync(turnContext, false, cancellationToken);

            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error at OnTurnAsync()");
                await base.OnTurnAsync(turnContext, cancellationToken);
            }
        }

        /// <summary>
        /// Invoked when a message activity is received from the user.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        /// <remarks>
        /// Reference link: https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.activityhandler.onmessageactivityasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        protected override async Task OnMessageActivityAsync(
            ITurnContext<IMessageActivity> turnContext,
            CancellationToken cancellationToken)
        {
            try
            {
                var message = turnContext?.Activity;
                this.telemetryClient.TrackEvent(
                    EVENT_MESSAGE_RECEIVED,
                    new Dictionary<string, string>
                    {
                        { "UserName" ,message.From.Name},
                        { "UserAadId", message.From?.AadObjectId ?? "" },
                        { "Product", options.ProductName },
                    });

                this.logger.LogInformation($"from: {message.From?.Id}, conversation: {message.Conversation.Id}, replyToId: {message.ReplyToId}");
                await this.SendTypingIndicatorAsync(turnContext).ConfigureAwait(false);

                switch (message.Conversation.ConversationType?.ToLower())
                {
                    case ConversationTypePersonal:
                        await this.OnMessageActivityInPersonalChatAsync(
                            message,
                            turnContext,
                            cancellationToken).ConfigureAwait(false);
                        break;

                    case ConversationTypeChannel:
                        await this.OnMessageActivityInChannelAsync(
                            message,
                            turnContext,
                            cancellationToken).ConfigureAwait(false);
                        break;
#if DEBUG
                    case null:
                        // in emulator treat all messages as personal chat
                        await this.OnMessageActivityInPersonalChatAsync(
                            message,
                            turnContext,
                            cancellationToken).ConfigureAwait(false);
                        break;
#endif
                    default:
                        this.logger.LogWarning($"Received unexpected conversationType {message.Conversation.ConversationType}");
                        break;
                }
                //await this.OnMessageActivityInPersonalChatAsync(
                //            message,
                //            turnContext,
                //            cancellationToken).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                await turnContext.SendActivityAsync(Strings.ErrorMessage).ConfigureAwait(false);
                this.logger.LogError(ex, $"Error processing message: {ex.Message}", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Invoke when a conversation update activity is received from the channel.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        /// <remarks>
        /// Reference link: https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.teams.teamsactivityhandler.onconversationupdateactivityasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        protected override async Task OnConversationUpdateActivityAsync(
            ITurnContext<IConversationUpdateActivity> turnContext,
            CancellationToken cancellationToken)
        {
            try
            {
                var activity = turnContext?.Activity;
                this.logger.LogInformation("Received conversationUpdate activity");
                this.logger.LogInformation($"conversationType: {activity.Conversation.ConversationType}, membersAdded: {activity.MembersAdded?.Count}, membersRemoved: {activity.MembersRemoved?.Count}");

                if (activity?.MembersAdded?.Count == 0)
                {
                    this.logger.LogInformation("Ignoring conversationUpdate that was not a membersAdded event");
                    return;
                }

                switch (activity.Conversation.ConversationType?.ToLower())
                {
                    case ConversationTypePersonal:
                        await this.OnMembersAddedToPersonalChatAsync(activity.MembersAdded, turnContext).ConfigureAwait(false);
                        return;

                    case ConversationTypeChannel:
                        await this.OnMembersAddedToTeamAsync(activity.MembersAdded, turnContext, cancellationToken).ConfigureAwait(false);
                        return;

                    default:
                        this.logger.LogInformation($"Ignoring event from conversation type {activity.Conversation.ConversationType}");
                        return;
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error processing conversationUpdate: {ex.Message}", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Invoke when user clicks on edit button on a question in SME team.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="taskModuleRequest">Task module invoke request value payload.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        /// <remarks>
        /// Reference link: https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.teams.teamsactivityhandler.onteamstaskmodulefetchasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        protected override async Task<TaskModuleResponse> OnTeamsTaskModuleFetchAsync(
            ITurnContext<IInvokeActivity> turnContext,
            TaskModuleRequest taskModuleRequest,
            CancellationToken cancellationToken)
        {
            try
            {
                var postedValues = JsonConvert.DeserializeObject<AdaptiveSubmitActionData>(JObject.Parse(taskModuleRequest?.Data?.ToString()).ToString());

                // if we are supploed just an ID, then load the qna from the Test environment database
                if (postedValues.QnaPairId != null && string.IsNullOrEmpty(postedValues.OriginalQuestion))
                {
                    // lookup the qnapaird
                    var knowledgeBaseId = await this.configurationProvider.GetSavedEntityDetailAsync(ConfigurationEntityTypes.KnowledgeBaseId).ConfigureAwait(false);
                    var qnaitems = await this.qnaServiceProvider.DownloadKnowledgebaseAsync(knowledgeBaseId, true);
                    var answerData = qnaitems.FirstOrDefault(x => x.Id == postedValues.QnaPairId.Value);

                    postedValues.OriginalQuestion = answerData.Questions[0];
                    postedValues.UpdatedQuestion = answerData.Questions[0];
                    if (Validators.IsValidJSON(answerData.Answer))
                    {
                        AnswerModel answerModel = JsonConvert.DeserializeObject<AnswerModel>(answerData.Answer);
                        postedValues.Description = answerModel.Description;
                        postedValues.Title = answerModel.Title;
                        postedValues.Subtitle = answerModel.Subtitle;
                        postedValues.ImageUrl = answerModel.ImageUrl;
                        postedValues.RedirectionUrl = answerModel.ImageUrl;
                    }
                    else
                    {
                        postedValues.Description = answerData.Answer;
                    }
                }

                var adaptiveCardEditor = MessagingExtensionQnaCard.AddQuestionForm(postedValues, this.appBaseUri);
                return await GetTaskModuleResponseAsync(adaptiveCardEditor, Strings.EditQuestionSubtitle, postedValues.QnaPairId);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while fetch event is received from the user.");
            }

            return default;
        }

        /// <summary>
        /// Invoked when the user submits a edited question from SME team.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="taskModuleRequest">Task module invoke request value payload.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        /// <remarks>
        /// Reference link: https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.teams.teamsactivityhandler.onteamstaskmodulesubmitasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        protected override async Task<TaskModuleResponse> OnTeamsTaskModuleSubmitAsync(
            ITurnContext<IInvokeActivity> turnContext,
            TaskModuleRequest taskModuleRequest,
            CancellationToken cancellationToken)
        {
            try
            {
                var postedQuestionData = ((JObject)turnContext?.Activity?.Value).GetValue("data", StringComparison.OrdinalIgnoreCase).ToObject<AdaptiveSubmitActionData>();
                if (postedQuestionData == null)
                {
                    await turnContext.SendActivityAsync(Strings.ErrorMessage).ConfigureAwait(false);
                    return default;
                }

                if (postedQuestionData.BackButtonCommandText == Strings.BackButtonCommandText)
                {
                    // Populates the prefilled data on task module for adaptive card form fields on back button click.
                    return await GetTaskModuleResponseAsync(MessagingExtensionQnaCard.AddQuestionForm(postedQuestionData, this.appBaseUri), Strings.EditQuestionSubtitle).ConfigureAwait(false);
                }

                if (postedQuestionData.PreviewButtonCommandText == Constants.PreviewCardCommandText)
                {
                    // Preview the actual view of the card on preview button click.
                    return await GetTaskModuleResponseAsync(MessagingExtensionQnaCard.PreviewCardResponse(postedQuestionData, this.appBaseUri)).ConfigureAwait(false);
                }

                return await this.RespondToQuestionTaskModuleAsync(postedQuestionData, turnContext).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                // Check if knowledge base is empty and has not published yet when sme user is trying to edit the qna pair.
                var errorResponseException = ex?.InnerException as ErrorResponseException;
                if (errorResponseException?.Response.StatusCode == HttpStatusCode.BadRequest)
                {
                    var knowledgeBaseId = await this.configurationProvider.GetSavedEntityDetailAsync(Constants.KnowledgeBaseEntityId).ConfigureAwait(false);
                    var hasPublished = await this.qnaServiceProvider.GetInitialPublishedStatusAsync(knowledgeBaseId).ConfigureAwait(false);

                    // Check if knowledge base has not published yet.
                    if (!hasPublished)
                    {
                        this.logger.LogError(ex, "Error while fetching the qna pair: knowledge base may be empty or it has not published yet.");
                        await turnContext.SendActivityAsync("Please wait for some time, updates to this question will be available in short time.").ConfigureAwait(false);
                    }
                }
                else
                {
                    this.logger.LogError(ex, "Error while submit event is received from the user.");
                    await turnContext.SendActivityAsync(Strings.ErrorMessage).ConfigureAwait(false);
                }
            }

            return default;
        }

        /// <summary>
        /// Invoked when the user opens the messaging extension or searching any content in it.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="query">Contains messaging extension query keywords.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>Messaging extension response object to fill compose extension section.</returns>
        /// <remarks>
        /// https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.teams.teamsactivityhandler.onteamsmessagingextensionqueryasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        protected override async Task<MessagingExtensionResponse> OnTeamsMessagingExtensionQueryAsync(
            ITurnContext<IInvokeActivity> turnContext,
            MessagingExtensionQuery query,
            CancellationToken cancellationToken)
        {
            var turnContextActivity = turnContext?.Activity;
            try
            {
                turnContextActivity.TryGetChannelData<TeamsChannelData>(out var teamsChannelData);
                string expertTeamId = await this.configurationProvider.GetSavedEntityDetailAsync(ConfigurationEntityTypes.TeamId).ConfigureAwait(false);

                if (turnContext != null && teamsChannelData?.Team?.Id == expertTeamId && await this.IsMemberOfSmeTeamAsync(turnContext).ConfigureAwait(false))
                {
                    var messageExtensionQuery = JsonConvert.DeserializeObject<MessagingExtensionQuery>(turnContextActivity.Value.ToString());
                    var searchQuery = this.GetSearchQueryString(messageExtensionQuery);

                    return new MessagingExtensionResponse
                    {
                        ComposeExtension = await SearchHelper.GetSearchResultAsync(searchQuery, messageExtensionQuery.CommandId, messageExtensionQuery.QueryOptions.Count, messageExtensionQuery.QueryOptions.Skip, turnContextActivity.LocalTimestamp, this.searchService, this.knowledgeBaseSearchService, this.activityStorageProvider).ConfigureAwait(false),
                    };
                }

                return new MessagingExtensionResponse
                {
                    ComposeExtension = new MessagingExtensionResult
                    {
                        Text = Strings.NonSMEErrorText,
                        Type = "message",
                    },
                };
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Failed to handle the messaging extension command {turnContextActivity.Name}: {ex.Message}", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Invoked when user clicks on "Add new question" button on messaging extension from SME team.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="action">Action to be performed.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>Response of messaging extension action.</returns>
        /// <remarks>
        /// Reference link: https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.teams.teamsactivityhandler.onteamsmessagingextensionfetchtaskasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        protected override Task<MessagingExtensionActionResponse> OnTeamsMessagingExtensionFetchTaskAsync(
            ITurnContext<IInvokeActivity> turnContext,
            MessagingExtensionAction action,
            CancellationToken cancellationToken)
        {
            try
            {
                turnContext.Activity.TryGetChannelData<TeamsChannelData>(out var teamsChannelData);
                string expertTeamId = this.configurationProvider.GetSavedEntityDetailAsync(ConfigurationEntityTypes.TeamId).GetAwaiter().GetResult();

                if (teamsChannelData?.Team?.Id != expertTeamId)
                {
                    var unauthorizedUserCard = MessagingExtensionQnaCard.UnauthorizedUserActionCard();
                    return Task.FromResult(new MessagingExtensionActionResponse
                    {
                        Task = new TaskModuleContinueResponse
                        {
                            Value = new TaskModuleTaskInfo
                            {
                                Card = unauthorizedUserCard,
                                Height = 250,
                                Width = 300,
                                Title = Strings.AddQuestionSubtitle,
                            },
                        },
                    });
                }

                var adaptiveCardEditor = MessagingExtensionQnaCard.AddQuestionForm(new AdaptiveSubmitActionData(), this.appBaseUri);
                return GetMessagingExtensionActionResponseAsync(adaptiveCardEditor, Strings.AddQuestionSubtitle);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while fetching task received by the bot.");
            }

            return default;
        }

        /// <summary>
        /// Invoked when the user submits a new question from SME team.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="action">Action to be performed.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>Response of messaging extension action.</returns>
        /// <remarks>
        /// Reference link: https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.teams.teamsactivityhandler.onteamsmessagingextensionsubmitactionasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        protected override async Task<MessagingExtensionActionResponse> OnTeamsMessagingExtensionSubmitActionAsync(
            ITurnContext<IInvokeActivity> turnContext,
            MessagingExtensionAction action,
            CancellationToken cancellationToken)
        {
            try
            {
                var postedQuestionObject = ((JObject)turnContext.Activity.Value).GetValue("data", StringComparison.OrdinalIgnoreCase).ToObject<AdaptiveSubmitActionData>();
                if (postedQuestionObject == null)
                {
                    return default;
                }

                if (postedQuestionObject.BackButtonCommandText == Strings.BackButtonCommandText)
                {
                    // Populates the prefilled data on task module for adaptive card form fields on back button click.
                    return await GetMessagingExtensionActionResponseAsync(MessagingExtensionQnaCard.AddQuestionForm(postedQuestionObject, this.appBaseUri), Strings.AddQuestionSubtitle).ConfigureAwait(false);
                }

                if (postedQuestionObject.PreviewButtonCommandText == Constants.PreviewCardCommandText)
                {
                    // Preview the actual view of the card on preview button click.
                    return await GetMessagingExtensionActionResponseAsync(MessagingExtensionQnaCard.PreviewCardResponse(postedQuestionObject, this.appBaseUri)).ConfigureAwait(false);
                }

                // Response of messaging extension action.
                return await this.RespondToQuestionMessagingExtensionAsync(postedQuestionObject, turnContext, cancellationToken).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                if (((ErrorResponseException)ex).Body?.Error?.Code == ErrorCodeType.QuotaExceeded)
                {
                    this.logger.LogError(ex, "QnA storage limit exceeded and is not able to save the qna pair. Please contact your system administrator to provision additional storage space.");
                    await turnContext.SendActivityAsync("QnA storage limit exceeded and is not able to save the qna pair. Please contact your system administrator to provision additional storage space.").ConfigureAwait(false);
                    return null;
                }

                this.logger.LogError(ex, "Error at OnTeamsMessagingExtensionSubmitActionAsync()");
                await turnContext.SendActivityAsync(Strings.ErrorMessage).ConfigureAwait(false);
            }

            return default;
        }

        /// <summary>
        /// Get TaskModuleResponse object while adding or editing the question and answer pair.
        /// </summary>
        /// <param name="questionAnswerAdaptiveCardEditor">Card as an input.</param>
        /// <param name="titleText">Gets or sets text that appears below the app name and to the right of the app icon.</param>
        /// <returns>Envelope for Task Module Response.</returns>
        private async Task<TaskModuleResponse> GetTaskModuleResponseAsync(Attachment questionAnswerAdaptiveCardEditor, string titleText = "", int? questionID = null)
        {
            string editFormUri = await this.configurationProvider.GetSavedEntityDetailAsync("EditFormUri");
            
            var taskModuleInfo = new TaskModuleTaskInfo
            {
                Height = TaskModuleHeight,
                Width = TaskModuleWidth,
                Title = titleText,
            };

            // use a rich edit form if available and we know which question it is, otherwise use adaptive card
            if (string.IsNullOrWhiteSpace(editFormUri) || questionID == null)
            {
                taskModuleInfo.Card = questionAnswerAdaptiveCardEditor;
            }
            else
            {
                Guid randomId = Guid.NewGuid();
                taskModuleInfo.Url = editFormUri + $"/{questionID}/?rand={randomId}";
            }

            return new TaskModuleResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = taskModuleInfo,
                },
            };
        }

        /// <summary>
        /// Get messaging extension response object.
        /// </summary>
        /// <param name="questionAnswerAdaptiveCardEditor">Card as an input.</param>
        /// <param name="titleText">Gets or sets text that appears below the app name and to the right of the app icon.</param>
        /// <returns>Response of messaging extension action object.</returns>
        private async Task<MessagingExtensionActionResponse> GetMessagingExtensionActionResponseAsync(
            Attachment questionAnswerAdaptiveCardEditor,
            string titleText = "")
        {
            string editFormUri = await this.configurationProvider.GetSavedEntityDetailAsync("EditFormUri");

            var taskModuleInfo = new TaskModuleTaskInfo
            {
                Height = TaskModuleHeight,
                Width = TaskModuleWidth,
                Title = titleText,
            };

            // use a rich edit form if available and we know which question it is, otherwise use adaptive card
            if (string.IsNullOrWhiteSpace(editFormUri))
            {
                taskModuleInfo.Card = questionAnswerAdaptiveCardEditor;
            }
            else
            {
                Guid randomId = Guid.NewGuid();
                taskModuleInfo.Url = editFormUri + $"/0/?rand={randomId}";
            }

            return new MessagingExtensionActionResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = taskModuleInfo,
                },
            };
        }

        /// <summary>
        /// Return normal card as response if only question and answer fields are filled while adding the QnA pair in the knowledgebase.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="qnaPairEntity">Qna pair entity that contains question and answer information.</param>
        /// <param name="isRichCard">Indicate whether it's a rich card or normal. While adding the qna pair,
        /// if sme user is providing the value for fields like: image url or title or subtitle or redirection url then it's a rich card otherwise it will be a normal card containing only question and answer. </param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>Response of messaging extension action object.</returns>
        private async Task<MessagingExtensionActionResponse> AddQuestionCardResponseAsync(
        ITurnContext<IInvokeActivity> turnContext,
        AdaptiveSubmitActionData qnaPairEntity,
        bool isRichCard,
        CancellationToken cancellationToken)
        {
            string combinedDescription = QnaHelper.BuildCombinedDescriptionAsync(qnaPairEntity);

            try
            {
                // Check if question exist in the production/test knowledgebase & exactly the same question.
                var hasQuestionExist = await this.qnaServiceProvider.QuestionExistsInKbAsync(qnaPairEntity.UpdatedQuestion).ConfigureAwait(false);

                // Question already exist in knowledgebase.
                if (hasQuestionExist)
                {
                    // Response with question already exist(in test knowledgebase).
                    // If edited question text is already exist in the test knowledgebase.
                    qnaPairEntity.IsQuestionAlreadyExists = true;

                    // Messaging extension response object.
                    return await GetMessagingExtensionActionResponseAsync(MessagingExtensionQnaCard.AddQuestionForm(qnaPairEntity, this.appBaseUri)).ConfigureAwait(false);
                }
            }
            catch (Exception ex)
            {
                // Check if exception is not related to empty kb then add the qna pair otherwise throw it.
                if (((ErrorResponseException)ex).Response.StatusCode == HttpStatusCode.BadRequest)
                {
                    var knowledgeBaseId = await this.configurationProvider.GetSavedEntityDetailAsync(Constants.KnowledgeBaseEntityId).ConfigureAwait(false);
                    var hasPublished = await this.qnaServiceProvider.GetInitialPublishedStatusAsync(knowledgeBaseId).ConfigureAwait(false);

                    // Check if knowledge base has not published yet.
                    // If knowledge base has published then throw the error otherwise contiue to add the question & answer pair.
                    if (hasPublished)
                    {
                        this.logger.LogError(ex, "Error while checking if the question exists in knowledge base.");
                        throw;
                    }
                }
            }

            // Save the question in the knowledgebase.
            var activityReferenceId = Guid.NewGuid().ToString();
            await this.qnaServiceProvider.AddQnaAsync(qnaPairEntity.UpdatedQuestion?.Trim(), combinedDescription, turnContext.Activity.From.AadObjectId, turnContext.Activity.Conversation.Id, activityReferenceId).ConfigureAwait(false);
            qnaPairEntity.IsTestKnowledgeBase = true;
            await SendNewQnAPairActivity(turnContext, qnaPairEntity, isRichCard, activityReferenceId, cancellationToken).ConfigureAwait(false);

            return default;
        }

        private async Task SendNewQnAPairActivity(ITurnContext turnContext, AdaptiveSubmitActionData qnaPairEntity, bool isRichCard, string activityReferenceId, CancellationToken cancellationToken)
        {
            ResourceResponse activityResponse;

            // Rich card as response.
            if (isRichCard)
            {
                qnaPairEntity.IsPreviewCard = false;
                activityResponse = await turnContext.SendActivityAsync(MessageFactory.Attachment(MessagingExtensionQnaCard.ShowRichCard(qnaPairEntity, turnContext.Activity.From.Name, Strings.EntryCreatedByText)), cancellationToken).ConfigureAwait(false);
            }
            else
            {
                // Normal card as response.
                activityResponse = await turnContext.SendActivityAsync(MessageFactory.Attachment(MessagingExtensionQnaCard.ShowNormalCard(qnaPairEntity, turnContext.Activity.From.Name, actionPerformed: Strings.EntryCreatedByText)), cancellationToken).ConfigureAwait(false);
            }

            this.telemetryClient.TrackEvent(
                EVENT_QUESTION_ADDED, new Dictionary<string, string>
                {
                    { "UserName" ,turnContext.Activity.From.Name},
                    { "UserAadId", turnContext.Activity.From?.AadObjectId ?? "" },
                    { "QuestionId", qnaPairEntity.QnaPairId?.ToString() ?? "" },
                    { "Question", qnaPairEntity.UpdatedQuestion },
                    { "Product", options.ProductName },
                });

            this.logger.LogInformation($"Question added by: {turnContext.Activity.From.AadObjectId}");
            ActivityEntity activityEntity = new ActivityEntity { ActivityId = activityResponse.Id, ActivityReferenceId = activityReferenceId };
            bool operationStatus = await this.activityStorageProvider.AddActivityEntityAsync(activityEntity).ConfigureAwait(false);
            if (!operationStatus)
            {
                this.logger.LogInformation($"Unable to add activity data in table storage.");
            }
        }

        /// <summary>
        /// Return card response.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="postedQnaPairEntity">Qna pair entity that contains question and answer information.</param>
        /// <param name="answer">Answer text.</param>
        /// <returns>Card attachment.</returns>
        private async Task<Attachment> CardResponseAsync(
            ITurnContext<IInvokeActivity> turnContext,
            AdaptiveSubmitActionData postedQnaPairEntity,
            string answer)
        {
            Attachment qnaAdaptiveCard = new Attachment();
            bool isSaved;

            if (postedQnaPairEntity.UpdatedQuestion?.ToUpperInvariant().Trim() == postedQnaPairEntity.OriginalQuestion?.ToUpperInvariant().Trim())
            {
                postedQnaPairEntity.IsTestKnowledgeBase = false;
                isSaved = await this.SaveQnaAsync(turnContext, answer, postedQnaPairEntity).ConfigureAwait(false);
                if (!isSaved)
                {
                    postedQnaPairEntity.IsTestKnowledgeBase = true;
                    await this.SaveQnaAsync(turnContext, answer, postedQnaPairEntity).ConfigureAwait(false);
                }
            }
            else
            {
                // Check if question exist in the production/test knowledgebase & exactly the same question.
                var hasQuestionExist = await this.qnaServiceProvider.QuestionExistsInKbAsync(postedQnaPairEntity.UpdatedQuestion).ConfigureAwait(false);

                // Edit the question if it doesn't exist in the test knowledgebse.
                if (hasQuestionExist)
                {
                    // If edited question text is already exist in the test knowledgebase.
                    postedQnaPairEntity.IsQuestionAlreadyExists = true;
                }
                else
                {
                    // Save the edited question in the knowledgebase.
                    postedQnaPairEntity.IsTestKnowledgeBase = false;
                    isSaved = await this.SaveQnaAsync(turnContext, answer, postedQnaPairEntity).ConfigureAwait(false);
                    if (!isSaved)
                    {
                        postedQnaPairEntity.IsTestKnowledgeBase = true;
                        await this.SaveQnaAsync(turnContext, answer, postedQnaPairEntity).ConfigureAwait(false);
                    }
                }

                if (postedQnaPairEntity.IsQuestionAlreadyExists)
                {
                    // Response with question already exist(in test knowledgebase).
                    qnaAdaptiveCard = MessagingExtensionQnaCard.AddQuestionForm(postedQnaPairEntity, this.appBaseUri);
                }
            }

            return qnaAdaptiveCard;
        }

        /// <summary>
        /// Handle 1:1 chat with members who started chat for the first time.
        /// </summary>
        /// <param name="membersAdded">Channel account information needed to route a message.</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task OnMembersAddedToPersonalChatAsync(
            IList<ChannelAccount> membersAdded,
            ITurnContext<IConversationUpdateActivity> turnContext)
        {
            var activity = turnContext.Activity;
            if (membersAdded.Any(channelAccount => channelAccount.Id == activity.Recipient.Id))
            {
                // User started chat with the bot in personal scope, for the first time.
                this.logger.LogInformation($"Bot added to 1:1 chat {activity.Conversation.Id}");
                var welcomeText = await this.configurationProvider.GetSavedEntityDetailAsync(ConfigurationEntityTypes.WelcomeMessageText).ConfigureAwait(false);
                var userWelcomeCardAttachment = WelcomeCard.GetCard(welcomeText);
                await turnContext.SendActivityAsync(MessageFactory.Attachment(userWelcomeCardAttachment)).ConfigureAwait(false);
            }
        }

        /// <summary>
        /// Handle members added conversationUpdate event in team.
        /// </summary>
        /// <param name="membersAdded">Channel account information needed to route a message.</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task OnMembersAddedToTeamAsync(
           IList<ChannelAccount> membersAdded,
           ITurnContext<IConversationUpdateActivity> turnContext,
           CancellationToken cancellationToken)
        {
            var activity = turnContext.Activity;
            if (membersAdded != null)
            {
                if (membersAdded.Any(channelAccount => channelAccount.Id == activity.Recipient.Id))
                {
                    // Bot was added to a team
                    this.logger.LogInformation($"Bot added to team {activity.Conversation.Id}");

                    var teamDetails = ((JObject)turnContext.Activity.ChannelData).ToObject<TeamsChannelData>();
                    var botDisplayName = turnContext.Activity.Recipient.Name;
                    var teamWelcomeCardAttachment = WelcomeTeamCard.GetCard();
                    await this.SendCardToTeamAsync(turnContext, teamWelcomeCardAttachment, teamDetails.Team.Id, cancellationToken).ConfigureAwait(false);
                }
            }
        }

        /// <summary>
        /// Handle message activity in 1:1 chat.
        /// </summary>
        /// <param name="message">A message in a conversation.</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task OnMessageActivityInPersonalChatAsync(
            IMessageActivity message,
            ITurnContext<IMessageActivity> turnContext,
            CancellationToken cancellationToken)
        {
            if (!string.IsNullOrEmpty(message.ReplyToId) && (message.Value != null) && ((JObject)message.Value).HasValues)
            {
                this.logger.LogInformation("Card submit in 1:1 chat");
                await this.OnAdaptiveCardSubmitInPersonalChatAsync(message, turnContext, cancellationToken).ConfigureAwait(false);
                return;
            }

            // Check if a file attachment of questions was received
            (string filename, IList<AnswerItem> questions, string language) = await TryGetQuestionsFromAttachment(turnContext.Activity);

            // If we have a questions from the attachment, process the them to find out the answers from qnA maker
            if (questions != null)
            {
                await ProcessInPersonalChatQuestionsAttachment(turnContext, filename, questions, language, cancellationToken);

            }
            else
            {
                // Process code as-is for non attachment
                await ProcessInPersonalChatMessage(message, turnContext);
            }
        }

        private async Task ProcessInPersonalChatQuestionsAttachment(ITurnContext<IMessageActivity> turnContext, string filename, IList<AnswerItem> questions, string questionLanguageCode, CancellationToken cancellationToken)
        {
            string extension = Path.GetExtension(filename);
            var sw = new System.Diagnostics.Stopwatch();

            await turnContext.SendActivityAsync($"I received your {extension.Substring(1)} file with {questions.Count} questions. One moment while I consult QnA Maker for the answers");
            await SendTypingIndicatorAsync(turnContext);

            IList<AnswerItem> normalizedQuestions;
            if (questionLanguageCode == translatorService.DefaultLanguageCode)
            {
                // just use the original list in place
                normalizedQuestions = questions;
            }
            else
            {
                sw.Start();
                // create a new set of questions in the default language
                // first translate all the questions at once
                var translationQuestions = await this.translatorService.TranslateAsync(
                    questions.Select(x => x.Question).ToArray(),
                    questionLanguageCode,
                    this.translatorService.DefaultLanguageCode
                );
                sw.Stop();
                this.telemetryClient.TrackEvent(
                    EVENT_TRANSLATED_QUESTION_BULK,
                    new Dictionary<string, string>
                    {
                        { "UserName" ,turnContext.Activity.From.Name},
                        { "UserAadId ", turnContext.Activity.From?.AadObjectId ?? "" },
                        { "Product", options.ProductName },
                        { "Language", questionLanguageCode },
                    },
                    new Dictionary<string, double>
                    {
                        { "Seconds", sw.Elapsed.TotalSeconds},
                        { "Count", questions.Count},
                    });

                // rebuild the normalizedQuestion list from the translated questions
                normalizedQuestions = new AnswerItem[questions.Count];
                for (int i = 0; i < questions.Count; i++)
                {
                    normalizedQuestions[i] = new AnswerItem();
                    normalizedQuestions[i].Metadata = questions[i].Metadata;
                    normalizedQuestions[i].Question = translationQuestions[i];
                }
            }

            sw.Start();

            // Create QnA object to store the answers from QnA
            for (int i = 0; i < normalizedQuestions.Count; i++)
            {
                var question = normalizedQuestions[i];
                if (string.IsNullOrEmpty(question.Question))
                {
                    question.Answer = "ERROR reading input";
                    question.Question = "ERROR reading input";
                }
                else
                {
                    try
                    {
                        QueryTag[] tags = question.Metadata?.Split('|').Select(x =>
                        {
                            x = x.Trim();
                            string[] parts = x.Split(':');
                            return new QueryTag() { Name = parts[0].Trim(), Value = parts[1].Trim() };
                        }).ToArray();

                        if (tags?.Length == 0)
                        {
                            tags = null;
                        }

                        question.Answer = await this.GenerateAnswer(question.Question, tags);
                    }
                    catch (Exception ex)
                    {
                        question.Answer = "ERROR generating answer";
                    }
                }

                if (i % 100 == 0 && i > 0)
                {
                    await turnContext.SendActivityAsync($"fyi - I've finished {i} so far");
                    await SendTypingIndicatorAsync(turnContext);
                }
            }

            sw.Stop();

            this.telemetryClient.TrackEvent(
                EVENT_ANSWERED_QUESTION_BULK,
                new Dictionary<string, string>
                {
                    { "UserName" ,turnContext.Activity.From.Name},
                    { "UserAadId ", turnContext.Activity.From?.AadObjectId ?? "" },
                    { "Product", options.ProductName },
                    { "Language", questionLanguageCode },
               },
                new Dictionary<string, double>
                {
                    { "Seconds", sw.Elapsed.TotalSeconds},
                    { "Count", questions.Count},
                });

            this.logger.LogInformation("Queried QnA Maker for {Count} questions in {Seconds} seconds", questions.Count, sw.Elapsed.TotalSeconds);

            if (questionLanguageCode != translatorService.DefaultLanguageCode)
            {
                sw.Start();
                // first translate all the answers at once
                var answersInOriginalLanguage = await this.translatorService.TranslateAsync(
                    normalizedQuestions.Select(x => x.Answer).ToArray(),
                    this.translatorService.DefaultLanguageCode,
                    questionLanguageCode
                );
                sw.Stop();
                this.telemetryClient.TrackEvent(
                    EVENT_TRANSLATED_ANSWERS_BULK,
                    new Dictionary<string, string>
                    {
                        { "UserName" ,turnContext.Activity.From.Name},
                        { "UserAadId ", turnContext.Activity.From?.AadObjectId ?? "" },
                        { "Product", options.ProductName },
                        { "Language", questionLanguageCode },
                    },
                    new Dictionary<string, double>
                    {
                        { "Seconds", sw.Elapsed.TotalSeconds},
                        { "Count", questions.Count},
                    });

                // populate the original question list with the translated answers
                for (int i = 0; i < questions.Count; i++)
                {
                    questions[i].Answer = answersInOriginalLanguage[i];
                }
            }

            string newFilename = Path.GetFileNameWithoutExtension(filename) + turnContext.Activity.Id + extension;

            byte[] answersContent = null;
            switch (extension)
            {
                case CSV_EXTENSION:
                    answersContent = CSVHelper.CsvFromQuestions(questions);
                    break;
                case XLSX_EXTENSION:
                    MemoryStream stream = new MemoryStream();
                    XlsxHelper.XlsxFromQuestions(questions, stream);
                    stream.Position = 0;
                    answersContent = stream.ToArray();
                    break;
            }

            await StoreBatchAnswers(turnContext.Activity.From.Id, answersContent);

            await this.SendFileCardAsync(turnContext, filename, newFilename, answersContent, cancellationToken);
        }

        /// <summary>
        /// store the results of a batch processing to table storage
        /// </summary>
        /// <param name="id">the key</param>
        /// <param name="answersContent">the content</param>
        /// <returns>a task</returns>
        private async Task StoreBatchAnswers(string id, byte[] answersContent)
        {
            await batchFileProvider.UpsertBatchFileAsync(new BatchFileEntity
            {
                Id = BatchResultsDirectory + id,
                FileBytes = answersContent
            });
        }

        /// <summary>
        /// retrieves the results of a batch processing to table storage
        /// </summary>
        /// <param name="id">the key</param>
        /// <returns>the file containing the answers</returns>
        private async Task<byte[]> GetBatchAnswers(string id)
        {
            var entity = await this.batchFileProvider.GetBatchFileAsync(BatchResultsDirectory + id);
            return entity?.FileBytes;
        }

        private async Task ProcessInPersonalChatMessage(IMessageActivity message, ITurnContext<IMessageActivity> turnContext)
        {
            // dialogs can happen in personal chat when its not an attachment
            await rootPersonalDialog.RunAsync(turnContext, conversationState.CreateProperty<DialogState>(nameof(DialogState)), CancellationToken.None);
        }

        private async Task<(string name, IList<AnswerItem> answers, string language)> TryGetQuestionsFromAttachment(IMessageActivity activity)
        {
            string contentUrl = null;
            string filename = null;
            string languageCode = translatorService.DefaultLanguageCode;

            IList<AnswerItem> answerItems = null;

            if (activity.ChannelId == Channels.Msteams)
            {
                // check for Teams attachments
                bool messageWithFileDownloadInfo = activity.Attachments?[0].ContentType == FileDownloadInfo.ContentType;
                var file = activity.Attachments?[0];
                filename = file?.Name;

                // check for attachments from Teams
                if (messageWithFileDownloadInfo)
                {
                    var fileDownload = JObject.FromObject(file.Content).ToObject<FileDownloadInfo>();
                    string filePath = Path.Combine("Files", filename);
                    contentUrl = fileDownload.DownloadUrl;
                }
            }
            else
            {
                // check for regular attachments (e.g. from the emulator)
                switch (activity.Attachments?[0].ContentType)
                {
                    case CSV_MIME_TYPE:
                        filename = activity.Attachments[0].Name;
                        contentUrl = activity.Attachments[0].ContentUrl;
                        break;
                    case XLSX_MIME_TYPE:
                        filename = activity.Attachments[0].Name;
                        contentUrl = activity.Attachments[0].ContentUrl;
                        break;
                }
            }

            // Read the file content
            if (contentUrl != null)
            {
                var client = clientFactory.CreateClient();
                var response = await client.GetAsync(contentUrl);

                switch (Path.GetExtension(filename))
                {
                    case CSV_EXTENSION:
                        string fileContent = await response.Content.ReadAsStringAsync();
                        answerItems = CSVHelper.AnswerListFromCsv(fileContent);
                        break;
                    case XLSX_EXTENSION:
                        using (Stream stream = await response.Content.ReadAsStreamAsync())
                        {
                            string tmpLanguageCode;
                            (answerItems, tmpLanguageCode) = XlsxHelper.QuestionsFromXlsx(stream);
                            if (!string.IsNullOrWhiteSpace(tmpLanguageCode) && await translatorService.IsValidTranslationLanguageCode(tmpLanguageCode))
                            {
                                languageCode = tmpLanguageCode;
                            }
                        }
                        break;
                }
            }

            return (filename, answerItems, languageCode);
        }

        /// <summary>
        /// Handle message activity in channel.
        /// </summary>
        /// <param name="message">A message in a conversation.</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task OnMessageActivityInChannelAsync(
            IMessageActivity message,
            ITurnContext<IMessageActivity> turnContext,
            CancellationToken cancellationToken)
        {
            string text;

            // Check if the incoming request is from SME for updating the ticket status.
            if (!string.IsNullOrEmpty(message.ReplyToId) && (message.Value != null) && ((JObject)message.Value).HasValues && !string.IsNullOrEmpty(((JObject)message.Value)["ticketId"]?.ToString()))
            {
                text = ChangeStatus;
            }
            else
            {
                text = message.Text?.ToLower()?.Trim() ?? string.Empty;
            }

            try
            {
                switch (text)
                {
                    case Constants.TeamTour:
                        this.logger.LogInformation("Sending team tour card");
                        var teamTourCards = TourCarousel.GetTeamTourCards(this.appBaseUri);
                        await turnContext.SendActivityAsync(MessageFactory.Carousel(teamTourCards)).ConfigureAwait(false);
                        break;

                    case ChangeStatus:
                        this.logger.LogInformation($"Card submit in channel {message.Value?.ToString()}");
                        await this.OnAdaptiveCardSubmitInChannelAsync(message, turnContext, cancellationToken).ConfigureAwait(false);
                        return;

                    case Constants.DeleteCommand:
                        this.logger.LogInformation($"Delete card submit in channel {message.Value?.ToString()}");
                        await QnaHelper.DeleteQnaPair(turnContext, this.qnaServiceProvider, this.activityStorageProvider, this.logger, cancellationToken).ConfigureAwait(false);
                        break;

                    case Constants.NoCommand:
                        return;

                    default:
                        this.logger.LogInformation("Unrecognized input in channel");
                        await turnContext.SendActivityAsync(MessageFactory.Attachment(UnrecognizedTeamInputCard.GetCard())).ConfigureAwait(false);
                        break;
                }
            }
            catch (Exception ex)
            {
                // Check if expert user is trying to delete the question and knowledge base has not published yet.
                if (((ErrorResponseException)ex).Response.StatusCode == HttpStatusCode.BadRequest)
                {
                    var knowledgeBaseId = await this.configurationProvider.GetSavedEntityDetailAsync(Constants.KnowledgeBaseEntityId).ConfigureAwait(false);
                    var hasPublished = await this.qnaServiceProvider.GetInitialPublishedStatusAsync(knowledgeBaseId).ConfigureAwait(false);

                    // Check if knowledge base has not published yet.
                    if (!hasPublished)
                    {
                        var activity = (Activity)turnContext.Activity;
                        var activityValue = ((JObject)activity.Value).ToObject<AdaptiveSubmitActionData>();
                        await turnContext.SendActivityAsync(MessageFactory.Text(string.Format(CultureInfo.InvariantCulture, Strings.WaitMessage, activityValue?.OriginalQuestion))).ConfigureAwait(false);
                        this.logger.LogError(ex, $"Error processing message: {ex.Message}", SeverityLevel.Error);
                        return;
                    }
                }

                // Throw the error at calling place, if there is any generic exception which is not caught by above conditon.
                throw;
            }
        }

        /// <summary>
        /// Handle adaptive card submit in 1:1 chat.
        /// Submits the question or feedback to the SME team.
        /// </summary>
        /// <param name="message">A message in a conversation.</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task OnAdaptiveCardSubmitInPersonalChatAsync(
            IMessageActivity message,
            ITurnContext<IMessageActivity> turnContext,
            CancellationToken cancellationToken)
        {
            Attachment smeTeamCard = null;      // Notification to SME team
            Attachment userCard = null;         // Acknowledgement to the user
            TicketEntity newTicket = null;      // New ticket

            switch (message?.Text)
            {
                case Constants.AskAnExpert:
                    this.logger.LogInformation("Sending user ask an expert card (from answer)");
                    var askAnExpertPayload = ((JObject)message.Value).ToObject<ResponseCardPayload>();
                    await turnContext.SendActivityAsync(MessageFactory.Attachment(AskAnExpertCard.GetCard(askAnExpertPayload))).ConfigureAwait(false);
                    break;

                case Constants.ShareFeedback:
                    this.logger.LogInformation("Sending user share feedback card (from answer)");
                    var shareFeedbackPayload = ((JObject)message.Value).ToObject<ResponseCardPayload>();
                    await turnContext.SendActivityAsync(MessageFactory.Attachment(ShareFeedbackCard.GetCard(shareFeedbackPayload))).ConfigureAwait(false);
                    break;

                case AskAnExpertCard.AskAnExpertSubmitText:
                    this.logger.LogInformation("Received question for expert");
                    newTicket = await AdaptiveCardHelper.AskAnExpertSubmitText(message, turnContext, cancellationToken, this.ticketsProvider).ConfigureAwait(false);
                    if (newTicket != null)
                    {
                        smeTeamCard = new SmeTicketCard(newTicket).ToAttachment(message?.LocalTimestamp);
                        userCard = new UserNotificationCard(newTicket).ToAttachment(Strings.NotificationCardContent, message?.LocalTimestamp);
                    }

                    break;

                case ShareFeedbackCard.ShareFeedbackSubmitText:
                    this.logger.LogInformation("Received app feedback");
                    smeTeamCard = await AdaptiveCardHelper.ShareFeedbackSubmitText(message, turnContext, cancellationToken).ConfigureAwait(false);
                    if (smeTeamCard != null)
                    {
                        await turnContext.SendActivityAsync(MessageFactory.Text(Strings.ThankYouTextContent)).ConfigureAwait(false);
                    }

                    break;

                default:
                    var payload = ((JObject)message.Value).ToObject<ResponseCardPayload>();

                    if (payload.IsPrompt)
                    {
                        this.logger.LogInformation("Sending input to QnAMaker for prompt");
                        await this.GetQuestionAnswerReplyAsync(turnContext, message).ConfigureAwait(false);
                    }
                    else
                    {
                        this.logger.LogWarning($"Unexpected text in submit payload: {message.Text}");
                    }

                    break;
            }

            string expertTeamId = await this.configurationProvider.GetSavedEntityDetailAsync(ConfigurationEntityTypes.TeamId).ConfigureAwait(false);

            // Send message to SME team.
            if (smeTeamCard != null)
            {
                var resourceResponse = await this.SendCardToTeamAsync(turnContext, smeTeamCard, expertTeamId, cancellationToken).ConfigureAwait(false);

                // If a ticket was created, update the ticket with the conversation info.
                if (newTicket != null)
                {
                    newTicket.SmeCardActivityId = resourceResponse.ActivityId;
                    newTicket.SmeThreadConversationId = resourceResponse.Id;
                    await this.ticketsProvider.UpsertTicketAsync(newTicket).ConfigureAwait(false);
                }
            }

            // Send acknowledgment to the user
            if (userCard != null)
            {
                await turnContext.SendActivityAsync(MessageFactory.Attachment(userCard), cancellationToken).ConfigureAwait(false);
            }
        }

        /// <summary>
        /// Handle adaptive card submit in channel.
        /// Updates the ticket status based on the user submission.
        /// </summary>
        /// <param name="message">A message in a conversation.</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task OnAdaptiveCardSubmitInChannelAsync(
            IMessageActivity message,
            ITurnContext<IMessageActivity> turnContext,
            CancellationToken cancellationToken)
        {
            var payload = ((JObject)message.Value).ToObject<ChangeTicketStatusPayload>();
            this.logger.LogInformation($"Received submit: ticketId={payload.TicketId} action={payload.Action}");

            // Get the ticket from the data store.
            var ticket = await this.ticketsProvider.GetTicketAsync(payload.TicketId).ConfigureAwait(false);
            if (ticket == null)
            {
                await turnContext.SendActivityAsync($"Ticket {payload.TicketId} was not found in the data store").ConfigureAwait(false);
                this.logger.LogInformation($"Ticket {payload.TicketId} was not found in the data store");
                return;
            }

            // Update the ticket based on the payload.
            switch (payload.Action)
            {
                case ChangeTicketStatusPayload.ReopenAction:
                    ticket.Status = (int)TicketState.Open;
                    ticket.DateAssigned = null;
                    ticket.AssignedToName = null;
                    ticket.AssignedToObjectId = null;
                    ticket.DateClosed = null;
                    break;

                case ChangeTicketStatusPayload.CloseAction:
                    ticket.Status = (int)TicketState.Closed;
                    ticket.DateClosed = DateTime.UtcNow;
                    break;

                case ChangeTicketStatusPayload.AssignToSelfAction:
                    ticket.Status = (int)TicketState.Open;
                    ticket.DateAssigned = DateTime.UtcNow;
                    ticket.AssignedToName = message.From.Name;
                    ticket.AssignedToObjectId = message.From.AadObjectId;
                    ticket.DateClosed = null;
                    break;

                default:
                    this.logger.LogWarning($"Unknown status command {payload.Action}");
                    return;
            }

            ticket.LastModifiedByName = message.From.Name;
            ticket.LastModifiedByObjectId = message.From.AadObjectId;
            await this.ticketsProvider.UpsertTicketAsync(ticket).ConfigureAwait(false);
            this.logger.LogInformation($"Ticket {ticket.TicketId} updated to status ({ticket.Status}, {ticket.AssignedToObjectId}) in store");

            // Update the card in the SME team.
            var updateCardActivity = new Activity(ActivityTypes.Message)
            {
                Id = ticket.SmeCardActivityId,
                Conversation = new ConversationAccount { Id = ticket.SmeThreadConversationId },
                Attachments = new List<Attachment> { new SmeTicketCard(ticket).ToAttachment(message.LocalTimestamp) },
            };
            var updateResponse = await turnContext.UpdateActivityAsync(updateCardActivity, cancellationToken).ConfigureAwait(false);
            this.logger.LogInformation($"Card for ticket {ticket.TicketId} updated to status ({ticket.Status}, {ticket.AssignedToObjectId}), activityId = {updateResponse.Id}");

            // Post update to user and SME team thread.
            string smeNotification = null;
            IMessageActivity userNotification = null;
            switch (payload.Action)
            {
                case ChangeTicketStatusPayload.ReopenAction:
                    smeNotification = string.Format(CultureInfo.InvariantCulture, Strings.SMEOpenedStatus, message.From.Name);

                    userNotification = MessageFactory.Attachment(new UserNotificationCard(ticket).ToAttachment(Strings.ReopenedTicketUserNotification, message.LocalTimestamp));
                    userNotification.Summary = Strings.ReopenedTicketUserNotification;
                    break;

                case ChangeTicketStatusPayload.CloseAction:
                    smeNotification = string.Format(CultureInfo.InvariantCulture, Strings.SMEClosedStatus, ticket.LastModifiedByName);

                    userNotification = MessageFactory.Attachment(new UserNotificationCard(ticket).ToAttachment(Strings.ClosedTicketUserNotification, message.LocalTimestamp));
                    userNotification.Summary = Strings.ClosedTicketUserNotification;
                    break;

                case ChangeTicketStatusPayload.AssignToSelfAction:
                    smeNotification = string.Format(CultureInfo.InvariantCulture, Strings.SMEAssignedStatus, ticket.AssignedToName);

                    userNotification = MessageFactory.Attachment(new UserNotificationCard(ticket).ToAttachment(Strings.AssignedTicketUserNotification, message.LocalTimestamp));
                    userNotification.Summary = Strings.AssignedTicketUserNotification;
                    break;
            }

            if (!string.IsNullOrEmpty(smeNotification))
            {
                var smeResponse = await turnContext.SendActivityAsync(smeNotification).ConfigureAwait(false);
                this.logger.LogInformation($"SME team notified of update to ticket {ticket.TicketId}, activityId = {smeResponse.Id}");
            }

            if (userNotification != null)
            {
                userNotification.Conversation = new ConversationAccount { Id = ticket.RequesterConversationId };
                var userResponse = await turnContext.Adapter.SendActivitiesAsync(turnContext, new Activity[] { (Activity)userNotification }, cancellationToken).ConfigureAwait(false);
                this.logger.LogInformation($"User notified of update to ticket {ticket.TicketId}, activityId = {userResponse.FirstOrDefault()?.Id}");
            }
        }

        /// <summary>
        /// Send typing indicator to the user.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task SendTypingIndicatorAsync(ITurnContext turnContext)
        {
            try
            {
                var typingActivity = turnContext.Activity.CreateReply();
                typingActivity.Type = ActivityTypes.Typing;
                await turnContext.SendActivityAsync(typingActivity);
            }
            catch (Exception ex)
            {
                // Do not fail on errors sending the typing indicator
                this.logger.LogWarning(ex, "Failed to send a typing indicator");
            }
        }

        /// <summary>
        /// Send the given attachment to the specified team.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cardToSend">The card to send.</param>
        /// <param name="teamId">Team id to which the message is being sent.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns><see cref="Task"/>That resolves to a <see cref="ConversationResourceResponse"/>Send a attachemnt.</returns>
        private async Task<ConversationResourceResponse> SendCardToTeamAsync(
            ITurnContext turnContext,
            Attachment cardToSend,
            string teamId,
            CancellationToken cancellationToken)
        {
            var conversationParameters = new ConversationParameters
            {
                Activity = (Activity)MessageFactory.Attachment(cardToSend),
                ChannelData = new TeamsChannelData { Channel = new ChannelInfo(teamId) },
            };

            var taskCompletionSource = new TaskCompletionSource<ConversationResourceResponse>();
            await ((BotFrameworkAdapter)turnContext.Adapter).CreateConversationAsync(
                null,       // If we set channel = "msteams", there is an error as preinstalled middleware expects ChannelData to be present.
                turnContext.Activity.ServiceUrl,
                this.microsoftAppCredentials,
                conversationParameters,
                (newTurnContext, newCancellationToken) =>
                {
                    var activity = newTurnContext.Activity;
                    taskCompletionSource.SetResult(new ConversationResourceResponse
                    {
                        Id = activity.Conversation.Id,
                        ActivityId = activity.Id,
                        ServiceUrl = activity.ServiceUrl,
                    });
                    return Task.CompletedTask;
                },
                cancellationToken).ConfigureAwait(false);

            return await taskCompletionSource.Task.ConfigureAwait(false);
        }

        /// <summary>
        /// Verify if the tenant Id in the message is the same tenant Id used when application was configured.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <returns>Boolean value where true represent tenant is valid while false represent tenant in not valid.</returns>
        private bool IsActivityFromExpectedTenant(ITurnContext turnContext)
        {
#if DEBUG
            // only check Tenant when in Teams
            if (turnContext.Activity.ChannelId != Channels.Msteams)
            {
                return true;
            }
            return true;
#endif
            return turnContext.Activity.Conversation.TenantId == this.options.TenantId;
        }

        /// <summary>
        /// Method perform update operation of question and answer pair.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="answer">Answer of the given question.</param>
        /// <param name="qnaPairEntity">Qna pair entity that contains question and answer information.</param>
        /// <returns>A <see cref="Task"/> of type bool where true represents question and answer pair updated successfully while false indicates failure in updating the question and answer pair.</returns>
        private async Task<bool> SaveQnaAsync(ITurnContext turnContext, string answer, AdaptiveSubmitActionData qnaPairEntity)
        {
            QnASearchResult searchResult;
            var qnaAnswerResponse = await this.qnaServiceProvider.GenerateAnswerAsync(qnaPairEntity.OriginalQuestion, qnaPairEntity.IsTestKnowledgeBase).ConfigureAwait(false);
            searchResult = qnaAnswerResponse.Answers.FirstOrDefault();
            bool isSameQuestion = false;

            // Check if question exist in the knowledgebase.
            if (searchResult != null && searchResult.Questions.Count > 0)
            {
                // Check if the edited question & result returned from the knowledgebase are same.
                isSameQuestion = searchResult.Questions.First().ToUpperInvariant() == qnaPairEntity.OriginalQuestion.ToUpperInvariant();
            }

            // Edit the QnA pair if the question is exist in the knowledgebase & exactly the same question on which we are performing the action.
            if (searchResult.Id != -1 && isSameQuestion)
            {
                int qnaPairId = searchResult.Id.Value;
                this.logger.LogInformation($"Question updated by: {turnContext.Activity.Conversation.AadObjectId}");
                this.telemetryClient.TrackEvent(
                    EVENT_UPDATED_QUESTION,
                    new Dictionary<string, string>
                    {
                        { "UserName" ,turnContext.Activity.From.Name},
                        { "UserAadId", turnContext.Activity.From?.AadObjectId ?? string.Empty },
                        { "Question", searchResult.Questions[0] },
                        { "QuestionId", searchResult.Id.ToString() },
                        { "Product", options.ProductName },
                    });
                Attachment attachment = new Attachment();
                if (qnaPairEntity.IsRichCard)
                {
                    qnaPairEntity.IsPreviewCard = false;
                    qnaPairEntity.IsTestKnowledgeBase = true;
                    attachment = MessagingExtensionQnaCard.ShowRichCard(qnaPairEntity, turnContext.Activity.From.Name, Strings.LastEditedText);
                }
                else
                {
                    qnaPairEntity.IsTestKnowledgeBase = true;
                    qnaPairEntity.Description = answer;
                    attachment = MessagingExtensionQnaCard.ShowNormalCard(qnaPairEntity, turnContext.Activity.From.Name, actionPerformed: Strings.LastEditedText);
                }

                string activityRefId = qnaAnswerResponse.Answers.First().Metadata.FirstOrDefault(x => x.Name == Constants.MetadataActivityReferenceId)?.Value;
                if (activityRefId != null)
                {
                    await this.qnaServiceProvider.UpdateQnaAsync(qnaPairId, answer, turnContext.Activity.From.AadObjectId, qnaPairEntity.UpdatedQuestion, qnaPairEntity.OriginalQuestion).ConfigureAwait(false);
                    IList<ActivityEntity> activityEntities = await activityStorageProvider.GetAsync(activityRefId);
                    var activityId = activityEntities.FirstOrDefault().ActivityId;
                    var updateCardActivity = new Activity(ActivityTypes.Message)
                    {
                        Id = activityId,
                        Conversation = turnContext.Activity.Conversation,
                        Attachments = new List<Attachment> { attachment },
                    };

                    // Send edited question and answer card as response.
                    await turnContext.UpdateActivityAsync(updateCardActivity, cancellationToken: default).ConfigureAwait(false);
                }
                else
                {
                    // this is a migrated kb so it doesnt have an activity yet.
                    // send the card as if it was newly created to a new conversation thread
                    string conversationID = turnContext.Activity.Conversation.Id;

                    // send to new conversation ID in teams - so strip off the thread ID if its there
                    conversationID = conversationID.Substring(0, conversationID.IndexOf(";"));
                    activityRefId = Guid.NewGuid().ToString();
                    turnContext.Activity.Conversation.Id = conversationID;
                    await this.qnaServiceProvider.UpdateQnaAsync(qnaPairId, answer, turnContext.Activity.From.AadObjectId, qnaPairEntity.UpdatedQuestion, qnaPairEntity.OriginalQuestion, conversationID, activityRefId);
                    await SendNewQnAPairActivity(turnContext, qnaPairEntity, true, activityRefId, cancellationToken: default);
                }
            }
            else
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// Get the value of the searchText parameter in the messaging extension query.
        /// </summary>
        /// <param name="query">Contains messaging extension query keywords.</param>
        /// <returns>A value of the searchText parameter.</returns>
        private string GetSearchQueryString(MessagingExtensionQuery query)
        {
            var messageExtensionInputText = query.Parameters.FirstOrDefault(parameter => parameter.Name.Equals(SearchTextParameterName, StringComparison.OrdinalIgnoreCase));
            return messageExtensionInputText?.Value?.ToString();
        }

        /// <summary>
        /// Check if user using the app is a valid SME or not.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <returns>A <see cref="Task"/> of type bool where true represents that user using the app is a valid SME while false indicates that user using the app is not a valid SME.</returns>
        private async Task<bool> IsMemberOfSmeTeamAsync(ITurnContext<IInvokeActivity> turnContext)
        {
            bool isUserPartOfRoster = false;
            string expertTeamId = await this.configurationProvider.GetSavedEntityDetailAsync(ConfigurationEntityTypes.TeamId).ConfigureAwait(false);
            try
            {
                ConversationAccount conversationAccount = new ConversationAccount()
                {
                    Id = expertTeamId,
                };

                ConversationReference conversationReference = new ConversationReference()
                {
                    ServiceUrl = turnContext.Activity.ServiceUrl,
                    Conversation = conversationAccount,
                };

                string currentUserId = turnContext.Activity.From.Id;

                // Check for current user id in cache and add id of current user to cache if they are not added before
                // once they are validated against SME roster.
                if (!this.accessCache.TryGetValue(currentUserId, out string membersCacheEntry))
                {
                    await this.botAdapter.ContinueConversationAsync(
                        this.appId,
                        conversationReference,
                        async (newTurnContext, newCancellationToken) =>
                        {
                            var members = await this.botAdapter.GetConversationMembersAsync(newTurnContext, default(CancellationToken)).ConfigureAwait(false);
                            foreach (var member in members)
                            {
                                if (member.Id == currentUserId)
                                {
                                    membersCacheEntry = member.Id;
                                    isUserPartOfRoster = true;
                                    var cacheEntryOptions = new MemoryCacheEntryOptions().SetSlidingExpiration(TimeSpan.FromDays(this.accessCacheExpiryInDays));
                                    this.accessCache.Set(currentUserId, membersCacheEntry, cacheEntryOptions);
                                    break;
                                }
                            }
                        },
                        default(CancellationToken)).ConfigureAwait(false);
                }
                else
                {
                    isUserPartOfRoster = true;
                }
            }
            catch (Exception error)
            {
                this.logger.LogError(error, $"Failed to get members of team {expertTeamId}: {error.Message}", SeverityLevel.Error);
                isUserPartOfRoster = false;
                throw;
            }

            return isUserPartOfRoster;
        }

        /// <summary>
        ///  Validate the adaptiver card fields while adding the question and answer pair.
        /// </summary>
        /// <param name="postedQnaPairEntity">Qna pair entity contains submitted card data.</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>Response of messaging extension action.</returns>
        private async Task<MessagingExtensionActionResponse> RespondToQuestionMessagingExtensionAsync(
            AdaptiveSubmitActionData postedQnaPairEntity,
            ITurnContext<IInvokeActivity> turnContext,
            CancellationToken cancellationToken)
        {
            // Check if fields contains Html tags or Question and answer empty then return response with error message.
            if (Validators.IsContainsHtml(postedQnaPairEntity) || Validators.IsQnaFieldsNullOrEmpty(postedQnaPairEntity))
            {
                // Returns the card with validation errors on add QnA task module.
                return await GetMessagingExtensionActionResponseAsync(MessagingExtensionQnaCard.AddQuestionForm(Validators.HtmlAndQnaEmptyValidation(postedQnaPairEntity), this.appBaseUri)).ConfigureAwait(false);
            }

            if (Validators.IsRichCard(postedQnaPairEntity))
            {
                // While adding the new entry in knowledgebase,if user has entered invalid Image URL or Redirect URL then show the error message to user.
                if (Validators.IsImageUrlInvalid(postedQnaPairEntity) || Validators.IsRedirectionUrlInvalid(postedQnaPairEntity))
                {
                    return await GetMessagingExtensionActionResponseAsync(MessagingExtensionQnaCard.AddQuestionForm(Validators.ValidateImageAndRedirectionUrls(postedQnaPairEntity), this.appBaseUri)).ConfigureAwait(false);
                }

                // Return the rich card as response to user if he has filled title & image URL while adding the new entry in knowledgebase.
                return await this.AddQuestionCardResponseAsync(turnContext, postedQnaPairEntity, isRichCard: true, cancellationToken).ConfigureAwait(false);
            }

            // Normal card as response if only question and answer fields are filled while adding the QnA pair in the knowledgebase.
            return await this.AddQuestionCardResponseAsync(turnContext, postedQnaPairEntity, isRichCard: false, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Validate the adaptive card fields while editing the question and answer pair.
        /// </summary>
        /// <param name="postedQnaPairEntity">Qna pair entity contains submitted card data.</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <returns>Envelope for Task Module Response.</returns>
        private async Task<TaskModuleResponse> RespondToQuestionTaskModuleAsync(
            AdaptiveSubmitActionData postedQnaPairEntity,
            ITurnContext<IInvokeActivity> turnContext)
        {
            // Check if fields contains Html tags or Question and answer empty then return response with error message.
            if (Validators.IsContainsHtml(postedQnaPairEntity) || Validators.IsQnaFieldsNullOrEmpty(postedQnaPairEntity))
            {
                // Returns the card with validation errors on add QnA task module.
                return await GetTaskModuleResponseAsync(MessagingExtensionQnaCard.AddQuestionForm(Validators.HtmlAndQnaEmptyValidation(postedQnaPairEntity), this.appBaseUri)).ConfigureAwait(false);
            }

            if (Validators.IsRichCard(postedQnaPairEntity))
            {
                if (Validators.IsImageUrlInvalid(postedQnaPairEntity) || Validators.IsRedirectionUrlInvalid(postedQnaPairEntity))
                {
                    // Show the error message on task module response for edit QnA pair, if user has entered invalid image or redirection url.
                    return await GetTaskModuleResponseAsync(MessagingExtensionQnaCard.AddQuestionForm(Validators.ValidateImageAndRedirectionUrls(postedQnaPairEntity), this.appBaseUri)).ConfigureAwait(false);
                }

                string combinedDescription = QnaHelper.BuildCombinedDescriptionAsync(postedQnaPairEntity);
                postedQnaPairEntity.IsRichCard = true;

                if (postedQnaPairEntity.UpdatedQuestion?.ToUpperInvariant().Trim() == postedQnaPairEntity.OriginalQuestion?.ToUpperInvariant().Trim())
                {
                    // Save the QnA pair, return the response and closes the task module.
                    await GetTaskModuleResponseAsync(this.CardResponseAsync(
                        turnContext,
                        postedQnaPairEntity,
                        combinedDescription).Result).ConfigureAwait(false);
                    return default;
                }
                else
                {
                    var hasQuestionExist = await this.qnaServiceProvider.QuestionExistsInKbAsync(postedQnaPairEntity.UpdatedQuestion).ConfigureAwait(false);
                    if (hasQuestionExist)
                    {
                        // Shows the error message on task module, if question already exist.
                        return await GetTaskModuleResponseAsync(this.CardResponseAsync(
                            turnContext,
                            postedQnaPairEntity,
                            combinedDescription).Result).ConfigureAwait(false);
                    }
                    else
                    {
                        // Save the QnA pair, return the response and closes the task module.
                        await GetTaskModuleResponseAsync(this.CardResponseAsync(
                            turnContext,
                            postedQnaPairEntity,
                            combinedDescription).Result).ConfigureAwait(false);
                        return default;
                    }
                }
            }
            else
            {
                // Normal card section.
                if (postedQnaPairEntity.UpdatedQuestion?.ToUpperInvariant().Trim() == postedQnaPairEntity.OriginalQuestion?.ToUpperInvariant().Trim())
                {
                    // Save the QnA pair, return the response and closes the task module.
                    await GetTaskModuleResponseAsync(this.CardResponseAsync(
                        turnContext,
                        postedQnaPairEntity,
                        postedQnaPairEntity.Description).Result).ConfigureAwait(false);
                    return default;
                }
                else
                {
                    var hasQuestionExist = await this.qnaServiceProvider.QuestionExistsInKbAsync(postedQnaPairEntity.UpdatedQuestion).ConfigureAwait(false);
                    if (hasQuestionExist)
                    {
                        // Shows the error message on task module, if question already exist.
                        return await GetTaskModuleResponseAsync(this.CardResponseAsync(
                            turnContext,
                            postedQnaPairEntity,
                            postedQnaPairEntity.Description).Result).ConfigureAwait(false);
                    }
                    else
                    {
                        // Save the QnA pair, return the response and closes the task module.
                        await GetTaskModuleResponseAsync(this.CardResponseAsync(
                            turnContext,
                            postedQnaPairEntity,
                            postedQnaPairEntity.Description).Result).ConfigureAwait(false);
                        return default;
                    }
                }
            }
        }

        /// <summary>
        /// Get the reply to a question asked by end user.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="message">Text message.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task GetQuestionAnswerReplyAsync(
            ITurnContext<IMessageActivity> turnContext,
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

                queryResult = await this.qnaServiceProvider.GenerateAnswerAsync(question: text, isTestKnowledgeBase: false, payload.PreviousQuestions?.First().Id.ToString(), payload.PreviousQuestions?.First().Questions.First()).ConfigureAwait(false);

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
                        await turnContext.SendActivityAsync(MessageFactory.Attachment(ResponseCard.GetCard(answerData, text, this.appBaseUri, payload))).ConfigureAwait(false);
                    }

                    this.telemetryClient.TrackEvent(
                        EVENT_ANSWERED_QUESTION_SINGLE,
                        new Dictionary<string, string>
                        {
                                { "QuestionId" ,payload.QnaPairId.ToString() },
                                { "QuestionAnswered", queryResult.Answers[0].Questions[0] },
                                { "QuestionAsked", text },
                                { "UserName" ,turnContext.Activity.From.Name},
                                { "UserAadId", turnContext.Activity.From?.AadObjectId ?? "" },
                                { "Product", options.ProductName },
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
                if (((ErrorResponseException)ex).Response.StatusCode == HttpStatusCode.BadRequest)
                {
                    var knowledgeBaseId = await this.configurationProvider.GetSavedEntityDetailAsync(Constants.KnowledgeBaseEntityId).ConfigureAwait(false);
                    var hasPublished = await this.qnaServiceProvider.GetInitialPublishedStatusAsync(knowledgeBaseId).ConfigureAwait(false);

                    // Check if knowledge base has not published yet.
                    if (!hasPublished)
                    {
                        this.logger.LogError(ex, "Error while fetching the qna pair: knowledge base may be empty or it has not published yet.");
                        await turnContext.SendActivityAsync(MessageFactory.Attachment(UnrecognizedInputCard.GetCard(text))).ConfigureAwait(false);
                        return;
                    }
                }

                // Throw the error at calling place, if there is any generic exception which is not caught.
                throw;
            }
        }

        private async Task<string> GenerateAnswer(string question, QueryTag[] tags = null)
        {
            var queryResult = await this.qnaServiceProvider.GenerateAnswerAsync(question: question, isTestKnowledgeBase: false, tags: tags).ConfigureAwait(false);
            var answer = string.Empty;

            if (queryResult.Answers.First().Id != -1)
            {
                var answerData = queryResult.Answers.First();
                answer = answerData.Answer;
            }
            return answer
                ;
        }

        private async Task SendFileCardAsync(ITurnContext turnContext, string questionFilename, string answerFilename, byte[] bytes, CancellationToken cancellationToken)
        {
            Activity replyActivity;
            string message = null;
            Attachment attachment = null;

            switch (turnContext.Activity.ChannelId)
            {
                case Channels.Msteams:
                    message = $"I now have the answers for your questions from QnA Maker. Please grant permission for me to upload them to your OneDrive. Thanks!";

                    var consentContext = new
                    {
                        filename = "filename",
                        id = turnContext.Activity.From.Id,
                    };

                    var fileCard = new FileConsentCard
                    {
                        Description = $"QnA Maker answers for \"{answerFilename}\"",
                        SizeInBytes = bytes.Length,
                        AcceptContext = consentContext,
                        DeclineContext = consentContext,
                    };

                    attachment = new Attachment
                    {
                        Content = fileCard,
                        ContentType = FileConsentCard.ContentType,
                        Name = answerFilename,
                    };

                    break;
#if DEBUG
                default:
                    message = $"I now have the answers for your questions from QnA Maker.";

                    string base64Data = Convert.ToBase64String(bytes);
                    string extension = Path.GetExtension(answerFilename);

                    string contentType;
                    switch (extension)
                    {
                        case CSV_EXTENSION:
                            contentType = CSV_MIME_TYPE;
                            break;
                        case XLSX_EXTENSION:
                            contentType = XLSX_MIME_TYPE;
                            break;
                        default:
                            throw new InvalidOperationException("Unknown file type");
                    }

                    attachment = new Attachment
                    {
                        Name = answerFilename,
                        ContentType = contentType,
                        ContentUrl = $"data:{contentType};base64,{base64Data}",
                    };
                    break;
#endif
            }

            replyActivity = turnContext.Activity.CreateReply(message);
            replyActivity.Attachments = new List<Attachment>() { attachment };
            await turnContext.SendActivityAsync(replyActivity, cancellationToken);
        }

        /// <summary>
        /// File Consent.
        /// </summary>
        /// <param name="turnContext">Turn Context.</param>
        /// <param name="fileConsentCardResponse">File Consent Card.</param>
        /// <param name="cancellationToken">Cancellation Token.</param>
        /// <returns>Async operation</returns>
        protected override async Task OnTeamsFileConsentAcceptAsync(ITurnContext<IInvokeActivity> turnContext, FileConsentCardResponse fileConsentCardResponse, CancellationToken cancellationToken)
        {
            try
            {
                this.logger.LogInformation("user accepted file consent");
                var context = JObject.FromObject(fileConsentCardResponse.Context);

                var client = this.clientFactory.CreateClient();
                byte[] answersContent = await GetBatchAnswers(context["id"].Value<string>());
                if (answersContent != null)
                {
                    this.logger.LogInformation("found file for user");
                    var content = new ByteArrayContent(answersContent);
                    content.Headers.ContentLength = answersContent.Length;
                    content.Headers.ContentRange = new ContentRangeHeaderValue(0, answersContent.Length - 1, answersContent.Length);
                    var response = await client.PutAsync(fileConsentCardResponse.UploadInfo.UploadUrl, content, cancellationToken);
                    response.EnsureSuccessStatusCode();
                    await this.FileUploadCompletedAsync(turnContext, fileConsentCardResponse, cancellationToken);
                    this.logger.LogInformation("uploaded file");
                }
                else
                {
                    this.logger.LogError("did not find file for user");
                }
            }
            catch (Exception e)
            {
                this.logger.LogError(e, "failed to upload file");
                await this.FileUploadFailedAsync(turnContext, e.ToString(), cancellationToken);
            }
        }

        /// <summary>
        /// File consent decline.
        /// </summary>
        /// <param name="turnContext">Turn Context.</param>
        /// <param name="fileConsentCardResponse">File Consent Card.</param>
        /// <param name="cancellationToken">Cancellation Token.</param>
        /// <returns>Async operation</returns>
        protected override async Task OnTeamsFileConsentDeclineAsync(ITurnContext<IInvokeActivity> turnContext, FileConsentCardResponse fileConsentCardResponse, CancellationToken cancellationToken)
        {
            JToken context = JObject.FromObject(fileConsentCardResponse.Context);

            var reply = MessageFactory.Text($"Declined. We won't upload file <b>{context["filename"]}</b>.");
            reply.TextFormat = "xml";
            await turnContext.SendActivityAsync(reply, cancellationToken);
        }

        /// <summary>
        /// File upload completed.
        /// </summary>
        /// <param name="turnContext">Turn Context.</param>
        /// <param name="fileConsentCardResponse">File Consent Card.</param>
        /// <param name="cancellationToken">Cancellation Token.</param>
        private async Task FileUploadCompletedAsync(ITurnContext turnContext, FileConsentCardResponse fileConsentCardResponse, CancellationToken cancellationToken)
        {
            var downloadCard = new FileInfoCard
            {
                UniqueId = fileConsentCardResponse.UploadInfo.UniqueId,
                FileType = fileConsentCardResponse.UploadInfo.FileType,
            };

            var asAttachment = new Attachment
            {
                Content = downloadCard,
                ContentType = FileInfoCard.ContentType,
                Name = fileConsentCardResponse.UploadInfo.Name,
                ContentUrl = fileConsentCardResponse.UploadInfo.ContentUrl,
            };

            var reply = MessageFactory.Text($"<b>File uploaded.</b> Your file <b>{fileConsentCardResponse.UploadInfo.Name}</b> is ready to download");
            reply.TextFormat = "xml";
            reply.Attachments = new List<Attachment> { asAttachment };

            await turnContext.SendActivityAsync(reply, cancellationToken);
        }

        private async Task FileUploadFailedAsync(ITurnContext turnContext, string error, CancellationToken cancellationToken)
        {
            var reply = MessageFactory.Text($"<b>File upload failed.</b> Error: <pre>{error}</pre>");
            reply.TextFormat = "xml";
            await turnContext.SendActivityAsync(reply, cancellationToken);
        }
    }
}