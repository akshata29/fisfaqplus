namespace Microsoft.Teams.Apps.FAQPlusPlus.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Schema;

    /// <summary>
    /// Middleware for translating text between the user and bot.
    /// Uses the Microsoft Translator Text API.
    /// </summary>
    public class TranslationMiddleware : IMiddleware
    {
        public const string PreferredLanguageSetting = "TranslationLanguagePreference";
        public const string PauseTranslationSetting = "TranslationLanguagePaused";
        private readonly TranslatorService translator;
        private readonly TranslationSettings translatorSettings;

        private readonly IStatePropertyAccessor<string> languageStateProperty;
        private readonly IStatePropertyAccessor<bool> pauseTranslationProperty;

        /// <summary>
        /// Initializes a new instance of the <see cref="TranslationMiddleware"/> class.
        /// </summary>
        /// <param name="translator">Translator implementation to be used for text translation.</param>
        /// <param name="translatorSettings">Default Language Settings</param>
        /// <param name="userState">User Parameter</param>
        public TranslationMiddleware(TranslatorService translator, TranslationSettings translatorSettings, UserState userState)
        {
            this.translator = translator ?? throw new ArgumentNullException(nameof(translator));
            this.translatorSettings = translatorSettings ?? throw new ArgumentNullException(nameof(translatorSettings));
            if (userState == null)
            {
                throw new ArgumentNullException(nameof(userState));
            }

            this.languageStateProperty = userState.CreateProperty<string>(PreferredLanguageSetting);
            this.pauseTranslationProperty = userState.CreateProperty<bool>(PauseTranslationSetting);
        }

        /// <summary>
        /// Processes an incoming activity.
        /// </summary>
        /// <param name="turnContext">Context object containing information for a single turn of conversation with a user.</param>
        /// <param name="next">The delegate to call to continue the bot middleware pipeline.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        public async Task OnTurnAsync(ITurnContext turnContext, NextDelegate next, CancellationToken cancellationToken = default(CancellationToken))
        {
            var defaultLanguage = this.translatorSettings.DefaultLanguage;

            if (turnContext == null)
            {
                throw new ArgumentNullException(nameof(turnContext));
            }

            (bool translate, string userLanguage) = await this.ShouldTranslateAsync(turnContext, cancellationToken);

            if (translate)
            {
                if (turnContext.Activity.Type == ActivityTypes.Message)
                {
                    turnContext.Activity.Text = await this.translator.TranslateAsync(turnContext.Activity.Text, defaultLanguage, cancellationToken);
                }

                turnContext.OnSendActivities(async (newContext, activities, nextSend) =>
                {
                    List<Task> tasks = new();
                    foreach (Activity currentActivity in activities.Where(a => a.Type == ActivityTypes.Message))
                    {
                        tasks.Add(this.TranslateMessageActivityAsync(currentActivity.AsMessageActivity(), userLanguage));
                    }

                    if (tasks.Any())
                    {
                        await Task.WhenAll(tasks).ConfigureAwait(false);
                    }

                    return await nextSend();
                });

                turnContext.OnUpdateActivity(async (newContext, activity, nextUpdate) =>
                {
                    // Translate messages sent to the user to user language
                    if (activity.Type == ActivityTypes.Message)
                    {
                        await this.TranslateMessageActivityAsync(activity.AsMessageActivity(), userLanguage);
                    }

                    return await nextUpdate();
                });
            }

            await next(cancellationToken).ConfigureAwait(false);
        }

        private async Task TranslateMessageActivityAsync(IMessageActivity activity, string targetLocale, CancellationToken cancellationToken = default(CancellationToken))
        {
            if (activity.Type == ActivityTypes.Message)
            {
                if (activity.Text is not null)
                {
                    activity.Text = await this.translator.TranslateAsync(activity.Text, targetLocale);
                }
            }
        }

        private async Task<(bool, string)> ShouldTranslateAsync(ITurnContext turnContext, CancellationToken cancellationToken = default(CancellationToken))
        {
            var defaultLanguage = this.translatorSettings.DefaultLanguage;

            try
            {
                string text = turnContext.Activity.Text;

                // dont translate if the bot has temporarily disabled it
                // e.g. when receiving the language preference from the user
                bool translationPaused = await pauseTranslationProperty.GetAsync(turnContext, () => false, cancellationToken);
                if (translationPaused)
                {
                    return (false, defaultLanguage);
                }

                // is the user's preferred language different from the default?
                string userLanguage = await this.languageStateProperty.GetAsync(turnContext, () => defaultLanguage, cancellationToken) ?? defaultLanguage;

                return (userLanguage != defaultLanguage, userLanguage);
            }
            catch (Exception ex)
            {
                return (false, defaultLanguage);
            }
        }
    }
}
