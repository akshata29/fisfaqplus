namespace Microsoft.Teams.Apps.FAQPlusPlus
{
    using System;
    using System.Net.Http;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Bot.Builder.TraceExtensions;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.FAQPlusPlus.Helpers;

    public class AdapterWithErrorHandler : BotFrameworkHttpAdapter
    {
        /// <summary> 
        /// Initializes a new instance of the <see cref="AdapterWithErrorHandler"/> class.
        /// </summary>
        /// <param name="configuration">Configuration</param>
        /// <param name="credentialProvider">Credential Providre</param>
        /// <param name="httpClient">Http Client</param>
        /// <param name="logger">logger</param>
        /// <param name="conversationState">conversationState</param>
        public AdapterWithErrorHandler(ICredentialProvider credentialProvider, IChannelProvider channelProvider, ConversationState conversationState, ILogger<BotFrameworkHttpAdapter> logger, TranslationMiddleware translationMiddleware)
            : base(credentialProvider, channelProvider, logger)
        {
            if (translationMiddleware != null)
            {
                // Add translation middleware to the adapter's middleware pipeline
                logger.LogInformation("Using translation middleware");
                this.Use(translationMiddleware);
            }

            OnTurnError = async (turnContext, exception) =>
            {
                // Log any leaked exception from the application.
                // NOTE: In production environment, you should consider logging this to
                // Azure Application Insights. Visit https://aka.ms/bottelemetry to see how
                // to add telemetry capture to your bot.
                logger.LogError(exception, $"[OnTurnError] unhandled error : {exception.Message}");

                // Send a message to the user
                await SendWithoutMiddleware(turnContext, "An unknown error occurred, please try again or if the error persists please notify the RFP CoE.");

                if (conversationState != null)
                {
                    try
                    {
                        // Delete the conversationState for the current conversation to prevent the
                        // bot from getting stuck in a error-loop caused by being in a bad state.
                        // ConversationState should be thought of as similar to "cookie-state" in a Web pages.
                        await conversationState.DeleteAsync(turnContext);
                    }
                    catch (Exception e)
                    {
                        logger.LogError(e, $"Exception caught on attempting to Delete ConversationState : {e.Message}");
                    }
                }

                // Send a trace activity, which will be displayed in the Bot Framework Emulator
                await turnContext.TraceActivityAsync("OnTurnError Trace", exception.Message, "https://www.botframework.com/schemas/error", "TurnError");
            };
        }

        private static async Task SendWithoutMiddleware(ITurnContext turnContext, string message)
        {
            // Sending the Activity directly through the Adapter rather than through the TurnContext skips the middleware processing
            // this might be important in this particular case because it might have been the TranslationMiddleware that is actually failing!
            var activity = MessageFactory.Text(message);

            // If we are skipping the TurnContext we must address the Activity manually here before sending it.
            activity.ApplyConversationReference(turnContext.Activity.GetConversationReference());

            // Send the actual Activity through the Adapter.
            await turnContext.Adapter.SendActivitiesAsync(turnContext, new[] { activity }, CancellationToken.None);
        }
    }
}
