using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.Dialogs.Choices;
using Microsoft.Extensions.Logging;
using Microsoft.Teams.Apps.FAQPlusPlus.Cards;
using Microsoft.Teams.Apps.FAQPlusPlus.Common;
using Microsoft.Teams.Apps.FAQPlusPlus.Helpers;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace Microsoft.Teams.Apps.FAQPlusPlus.Dialogs
{
    public class ChangeLanguageDialog : ComponentDialog
    {
        private readonly ILogger<ChangeLanguageDialog> _logger;
        private readonly TranslatorService _translator;
        private readonly IStatePropertyAccessor<string> _languagePreferenceProperty;

        public ChangeLanguageDialog(UserState userState, ILogger<ChangeLanguageDialog> logger, TranslatorService translator)
            : base(nameof(ChangeLanguageDialog))
        {
            _languagePreferenceProperty = userState.CreateProperty<string>(TranslationMiddleware.PreferredLanguageSetting);
            _logger = logger;
            _translator = translator;

            AddDialog(new TextPrompt("languagePrompt"));
            AddDialog(new WaterfallDialog("localRoot",
                new WaterfallStep[] { StartDialogAsync, ProcessLanguageResponse }));
            InitialDialogId = "localRoot";
        }

        private async Task<DialogTurnResult> StartDialogAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            string langPref = await _languagePreferenceProperty.GetAsync(stepContext.Context, () => _translator.DefaultLanguage, cancellationToken);
            return await stepContext.PromptAsync("languagePrompt", new PromptOptions
            {
                Prompt = MessageFactory.Text($"Your current language preference is '{langPref}' what would you like to change it to? (e.g. en, es, de)"),
            });
        }

        private async Task<DialogTurnResult> ProcessLanguageResponse(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            string choice = stepContext.Result.ToString()?.Trim()?.ToLower();

            if (choice != null)
            {
                if (await _translator.IsValidTranslationLanguage(choice))
                {
                    await stepContext.Context.SendActivityAsync($"OK, I'll set your language preference to '{choice}'.");
                    await _languagePreferenceProperty.SetAsync(stepContext.Context, choice);
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

            return await stepContext.EndDialogAsync(null, cancellationToken);
        }
    }
}
