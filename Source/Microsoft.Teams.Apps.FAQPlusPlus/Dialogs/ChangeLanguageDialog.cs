using Microsoft.ApplicationInsights;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.Dialogs.Choices;
using Microsoft.Extensions.Options;
using Microsoft.Teams.Apps.FAQPlusPlus.Bots;
using Microsoft.Teams.Apps.FAQPlusPlus.Common.Models.Configuration;
using Microsoft.Teams.Apps.FAQPlusPlus.Helpers;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace Microsoft.Teams.Apps.FAQPlusPlus.Dialogs
{
    public class ChangeLanguageDialog : ComponentDialog
    {
        private readonly TranslatorService _translator;
        private readonly IStatePropertyAccessor<string> _languagePreferenceProperty;
        private readonly IStatePropertyAccessor<bool> _pauseTranslationProperty;
        private readonly TelemetryClient _telemetryClient;
        private readonly BotSettings _options;

        public ChangeLanguageDialog(
                        UserState userState, 
                        TelemetryClient telemetryClient, 
                        TranslatorService translator, 
                        IOptionsMonitor<BotSettings> optionsAccessor
            )
            : base(nameof(ChangeLanguageDialog))
        {
            _languagePreferenceProperty = userState.CreateProperty<string>(TranslationMiddleware.PreferredLanguageSetting);
            _pauseTranslationProperty = userState.CreateProperty<bool>(TranslationMiddleware.PauseTranslationSetting);
            _translator = translator;
            _telemetryClient = telemetryClient;
            _options = optionsAccessor.CurrentValue;

            AddDialog(new ChoicePrompt("languagePrompt"));
            AddDialog(new WaterfallDialog("localRoot",
                new WaterfallStep[] { StartDialogAsync, ProcessLanguageResponse }));
            InitialDialogId = "localRoot";
        }

        private async Task<DialogTurnResult> StartDialogAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            // disable translation for this prompt because the choices should not be translated
            await _pauseTranslationProperty.SetAsync(stepContext.Context, true);

            string langPref = await _languagePreferenceProperty.GetAsync(stepContext.Context, () => _translator.DefaultLanguageCode, cancellationToken);

            string langPrefName = (await _translator.GetAvailableLanguages())[langPref].nativeName;
            return await stepContext.PromptAsync("languagePrompt", new PromptOptions
            {
                Style = ListStyle.HeroCard,
                RetryPrompt = MessageFactory.Text("I'm sorry I didn't understand that. Please try rephrasing"),
                Choices = new[] {
                    new Choice() { Value = "English" },
                    new Choice() { Value = "German", Synonyms = new List<string> { "Deutsch" } },
                    new Choice() { Value = "French", Synonyms = new List<string> { "Français", "francais"} },
                    new Choice() { Value = "Spanish", Synonyms = new List<string> { "Español", "espanol" } },
                    new Choice() { Value = "Japanese", Synonyms = new List<string> { "日本語" } },
                    new Choice() { Value = "Korean", Synonyms = new List<string> { "한국어" } },
                    new Choice() { Value = "Vietnamese", Synonyms = new List<string> { "Tiếng Việt", "tieng viet" } },
                    new Choice() { Value = "Chinese Simplified", Synonyms = new List<string> { "簡體中文", "jiǎn tǐ zhōng wén", "jian ti zhong wen" } },
                    new Choice() { Value = "Chinese Traditional", Synonyms = new List<string> { "繁體中文", "fán tǐ zhōng wén", "fan ti zhong wen" } },

                },
                Prompt = MessageFactory.Text($"Your current language preference is '{langPrefName}' what would you like to change it to?"),
            });
        }

        private async Task<DialogTurnResult> ProcessLanguageResponse(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            try
            {
                // re-enable translation
                await _pauseTranslationProperty.SetAsync(stepContext.Context, false);

                FoundChoice selection = stepContext.Result as FoundChoice;
                string choice = selection.Value.ToString()?.Trim();

                if (choice != null)
                {
                    if (await _translator.IsValidTranslationLanguage(choice))
                    {
                        await stepContext.Context.SendActivityAsync($"OK, I'll set your language preference to '{choice}'.");
                        string code = await _translator.GetCodeForLanguage(choice);
                        await _languagePreferenceProperty.SetAsync(stepContext.Context, code);

                        _telemetryClient.TrackEvent(
                            FaqPlusPlusBot.EVENT_LANGUAGE_PREFERENCE_CHANGED,
                            new Dictionary<string, string>
                            {
                                                { "UserName" ,stepContext.Context.Activity.From.Name},
                                                { "UserAadId ", stepContext.Context.Activity.From?.AadObjectId ?? "" },
                                                { "Product", _options.ProductName },
                                                { "Language", code },
                            });
                    }
                    else
                    {
                        await stepContext.Context.SendActivityAsync($"I'm sorry, '{choice}' is not a valid translation language");
                    }
                }
                else
                {
                    await stepContext.Context.SendActivityAsync($"I'm sorry, something went wrong, please try again later");
                }
            }
            catch (Exception ex)
            {
                await stepContext.Context.SendActivityAsync($"I'm sorry, something went wrong, please try again later");
                _telemetryClient.TrackException(ex,
                                                new Dictionary<string, string>
                                                {
                                                    { "UserName" ,stepContext.Context.Activity.From.Name},
                                                    { "UserAadId ", stepContext.Context.Activity.From?.AadObjectId ?? "" },
                                                    { "Product", _options.ProductName },
                                                    { "Message", "failed when attempting to set language preference" },
                                                });
            }

            return await stepContext.EndDialogAsync(null, cancellationToken);
        }
    }
}
