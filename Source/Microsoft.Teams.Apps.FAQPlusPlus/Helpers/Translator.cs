namespace Microsoft.Teams.Apps.FAQPlusPlus.Helpers
{
    using System;
    using System.Linq;
    using System.Net.Http;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Teams.Apps.FAQPlusPlus.Models;
    using Newtonsoft.Json;

    public class Translator
    {
        public readonly string DefaultLanguage = "en";

        private const string Host = "https://api.cognitive.microsofttranslator.com";
        private const string Path = "/translate?api-version=3.0";
        private const string UriParams = "&to=";

        private static readonly HttpClient client = new HttpClient();

        private readonly string key;
        private readonly string region;

        /// <summary>
        /// Initializes a new instance of the <see cref="Translator"/> class.
        /// </summary>
        /// <param name="configuration">Configuration info</param>
        public Translator(IConfiguration configuration)
        {
            this.key = configuration["TranslatorKey"] ?? throw new ArgumentNullException(nameof(key));
            this.region = configuration["TranslatorKeyRegion"] ?? throw new ArgumentNullException(nameof(region));
        }

        /// <summary>
        /// Translates an array of strings
        /// </summary>
        /// <param name="texts">The text strings to translate. Can be at most 100 strings, cannot exceed 10k chars including spaces</param>
        /// <param name="sourceLocale">the locale of <paramref name="texts"/></param>
        /// <param name="targetLocale">the locare to translate to</param>
        /// <param name="cancellationToken">a cancellation token</param>
        /// <returns>the translated strings</returns>
        public async Task<string[]> TranslateAsync(string[] texts, string sourceLocale, string targetLocale, CancellationToken cancellationToken = default(CancellationToken))
        {
            // TODO: ensure 1000 questions works
            // TODO: add telemetry
            var body = texts.Select(x => new { Text = x }).ToArray();
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

                return result.Select(x => x.Translations.First().Text).ToArray();
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

        private readonly string[] validLanguages = new string[] { "en", "es" };

        internal bool IsValidTranslationLanguage(string language)
        {
            return validLanguages.Contains(language);
        }
    }
}
