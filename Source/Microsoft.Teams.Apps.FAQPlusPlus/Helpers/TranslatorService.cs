namespace Microsoft.Teams.Apps.FAQPlusPlus.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net.Http;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.FAQPlusPlus.Models;
    using Newtonsoft.Json;

    public class TranslatorService
    {
        public readonly string DefaultLanguage = "en";

        private const string Host = "https://api.cognitive.microsofttranslator.com";
        private const string Path = "/translate?api-version=3.0";
        private const string UriParams = "&to=";
        private static readonly HttpClient client = new HttpClient();

        private readonly string key;
        private readonly string region;
        private readonly ILogger<TranslatorService> _logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="TranslatorService"/> class.
        /// </summary>
        /// <param name="configuration">Configuration info</param>
        public TranslatorService(IConfiguration configuration, ILogger<TranslatorService> logger)
        {
            this.key = configuration["TranslatorKey"] ?? throw new ArgumentNullException(nameof(key));
            this.region = configuration["TranslatorKeyRegion"] ?? throw new ArgumentNullException(nameof(region));
            _logger = logger;
        }

        /// <summary>
        /// Translates an array of strings
        /// </summary>
        /// <param name="texts">The text strings to translate. Can be at most 100 strings, cannot exceed 10k chars including spaces</param>
        /// <param name="sourceLocale">the locale of <paramref name="texts"/></param>
        /// <param name="targetLocale">the locare to translate to</param>
        /// <param name="cancellationToken">a cancellation token</param>
        /// <returns>the translated strings</returns>
        public async Task<IList<string>> TranslateAsync(string[] texts, string sourceLocale, string targetLocale, CancellationToken cancellationToken = default(CancellationToken))
        {
            // chunk texts into request blocks < 100 strings and < 10k char total

            List<List<string>> batches = new List<List<string>>();
            int charCount = 0;
            const int maxCharCount = 10000;
            const int maxStringCount = 100;
            int stringCount = 0;

            List<string> currentBatch = new List<string>();
            foreach (string text in texts)
            {
                if ((text.Length + charCount >= maxCharCount) || stringCount >= maxStringCount)
                {
                    batches.Add(currentBatch);
                    currentBatch = new List<string>();
                    stringCount = 0;
                    charCount = 0;
                }

                currentBatch.Add(text);
                stringCount++;
                charCount += text.Length;
            }

            batches.Add(currentBatch);

            List<string> results = new List<string>(texts.Length);
            foreach (var batch in batches)
            {
                var result = await TranslateStrings(batch);
                results.AddRange(result);
            }

            return results;

            async Task<IEnumerable<string>> TranslateStrings(IList<string> strings)
            {
                var body = strings.Select(x => new { Text = x }).ToArray();
                var requestBody = JsonConvert.SerializeObject(body);

                using (var request = new HttpRequestMessage())
                {
                    var uri = $"{Host}{Path}&to={targetLocale}&from={sourceLocale}";
                    request.Method = HttpMethod.Post;
                    request.RequestUri = new Uri(uri);
                    request.Content = new StringContent(requestBody, Encoding.UTF8, "application/json");
                    request.Headers.Add("Ocp-Apim-Subscription-Key", key);
                    request.Headers.Add("Ocp-Apim-Subscription-Region", region);

                    var response = await client.SendAsync(request, cancellationToken);

                    if (!response.IsSuccessStatusCode)
                    {
                        throw new Exception($"The call to the translation service returned HTTP status code {response.StatusCode}.");
                    }

                    var responseBody = await response.Content.ReadAsStringAsync();
                    var result = JsonConvert.DeserializeObject<TranslatorResponse[]>(responseBody);

                    return result.Select(x => x.Translations.First().Text);
                }
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="text">Text to Translate</param>
        /// <param name="targetLocale">Locale</param>
        /// <param name="cancellationToken">Cancellation Token</param>
        /// <returns>A <see cref="Task{TResult}"/> representing the result of the asynchronous operation.</returns>
        public async Task<string> TranslateAsync(string text, string targetLocale, CancellationToken cancellationToken = default(CancellationToken))
        {
            // From Cognitive Services translation documentation:
            // https://docs.microsoft.com/en-us/azure/cognitive-services/translator/quickstart-csharp-translate
            var body = new object[] { new { Text = text } };
            var requestBody = JsonConvert.SerializeObject(body);

            using (var request = new HttpRequestMessage())
            {
                var uri = Host + Path + UriParams + targetLocale;
                request.Method = HttpMethod.Post;
                request.RequestUri = new Uri(uri);
                request.Content = new StringContent(requestBody, Encoding.UTF8, "application/json");
                request.Headers.Add("Ocp-Apim-Subscription-Key", key);
                request.Headers.Add("Ocp-Apim-Subscription-Region", region);

                var response = await client.SendAsync(request, cancellationToken);

                if (!response.IsSuccessStatusCode)
                {
                    throw new Exception($"The call to the translation service returned HTTP status code {response.StatusCode}.");
                }

                var responseBody = await response.Content.ReadAsStringAsync();
                var result = JsonConvert.DeserializeObject<TranslatorResponse[]>(responseBody);

                return result?.FirstOrDefault()?.Translations?.FirstOrDefault()?.Text;
            }
        }

        private static Dictionary<string, TranslatorLanguageDescription> validLanguages = null;

        public async Task<IReadOnlyDictionary<string, TranslatorLanguageDescription>> GetAvailableLanguages()
        {
            if (validLanguages == null)
            {
                try
                {
                    string responseString = await client.GetStringAsync("https://api.cognitive.microsofttranslator.com/languages?api-version=3.0&scope=translation");
                    var response = JsonConvert.DeserializeObject<TranslatorLanguagesResponse>(responseString);
                    validLanguages = response.translation;
                }
                catch(Exception ex)
                {
                    _logger.LogError(ex, "Failed to get list of valid languages from translator service");
                }
            }

            return validLanguages;
        }

        public async Task<bool> IsValidTranslationLanguage(string language)
        {
            var validLanguages = await GetAvailableLanguages();
            return validLanguages.ContainsKey(language);
        }
    }
}
