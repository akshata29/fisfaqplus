// <copyright file="Startup.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.FAQPlusPlus
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.AspNetCore.Builder;
    using Microsoft.AspNetCore.Hosting;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Azure.CognitiveServices.Knowledge.QnAMaker;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.BotFramework;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Extensions.Caching.Memory;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.DependencyInjection;
    using Microsoft.Extensions.Hosting;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.FAQPlusPlus.Bots;
    using Microsoft.Teams.Apps.FAQPlusPlus.Common.Models.Configuration;
    using Microsoft.Teams.Apps.FAQPlusPlus.Common.Providers;
    using Microsoft.Teams.Apps.FAQPlusPlus.Common.Providers.QuestionAnswering;
    using Microsoft.Teams.Apps.FAQPlusPlus.Dialogs;
    using Microsoft.Teams.Apps.FAQPlusPlus.Helpers;

    /// <summary>
    /// This a Startup class for this Bot.
    /// </summary>
    public class Startup
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="Startup"/> class.
        /// </summary>
        /// <param name="configuration">Startup Configuration.</param>
        public Startup(IConfiguration configuration, IWebHostEnvironment environment)
        {
            this.Configuration = configuration;
            Environment = environment;
        }

        /// <summary>
        /// Gets Configurations Interfaces.
        /// </summary>
        public IConfiguration Configuration { get; }
        public IWebHostEnvironment Environment { get; }

        /// <summary>
        /// This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        /// </summary>
        /// <param name="app">Application Builder.</param>
        /// <param name="env">Hosting Environment.</param>
        public void Configure(IApplicationBuilder app)
        {
            if (Environment.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }
            else
            {
                app.UseHsts();
            }

            app.UseStaticFiles();
            app.UseRouting();
            app.UseDefaultFiles();
            app.UseEndpoints(endpoints => { endpoints.MapControllers(); });
        }

        private static string StripRouteFromQnAMakerEndpoint(string endpoint)
        {
            const string apiRoute = "/qnamaker/v5.0-preview.1";

            if (endpoint.EndsWith(apiRoute, System.StringComparison.OrdinalIgnoreCase))
            {
                endpoint = endpoint.Substring(0, endpoint.Length - apiRoute.Length);
            }
            return endpoint;
        }

        /// <summary>
        /// This method gets called by the runtime. Use this method to add services to the container.
        /// </summary>
        /// <param name="services"> Service Collection Interface.</param>
        public void ConfigureServices(IServiceCollection services)
        {
            services.AddApplicationInsightsTelemetry();
            services.Configure<KnowledgeBaseSettings>(knowledgeBaseSettings =>
            {
                knowledgeBaseSettings.SearchServiceName = this.Configuration["SearchServiceName"];
                knowledgeBaseSettings.SearchServiceQueryApiKey = this.Configuration["SearchServiceQueryApiKey"];
                knowledgeBaseSettings.SearchServiceAdminApiKey = this.Configuration["SearchServiceAdminApiKey"];
                knowledgeBaseSettings.SearchIndexingIntervalInMinutes = this.Configuration["SearchIndexingIntervalInMinutes"];
                knowledgeBaseSettings.StorageConnectionString = this.Configuration["StorageConnectionString"];
            });

            services.Configure<QnAMakerSettings>(qnAMakerSettings =>
            {
                qnAMakerSettings.ScoreThreshold = this.Configuration["ScoreThreshold"];
            });

            services.Configure<BotSettings>(botSettings =>
            {
                botSettings.AccessCacheExpiryInDays = Convert.ToInt32(this.Configuration["AccessCacheExpiryInDays"]);
                botSettings.AppBaseUri = this.Configuration["AppBaseUri"];
                botSettings.MicrosoftAppId = this.Configuration["MicrosoftAppId"];
                botSettings.TenantId = this.Configuration["TenantId"];
                string productName = this.Configuration["ProductName"];
                if (!string.IsNullOrEmpty(productName))
                {
                    botSettings.ProductName = productName;
                }
                else
                {
                    botSettings.ProductName = "Not Set";
                }
            });

            services.AddMvc().SetCompatibilityVersion(CompatibilityVersion.Version_2_1);
            services.AddSingleton<Common.Providers.IConfigurationDataProvider>(new Common.Providers.ConfigurationDataProvider(this.Configuration["StorageConnectionString"]));
            services.AddHttpClient();

            // Configure channel provider
            services.AddSingleton<IChannelProvider, ConfigurationChannelProvider>();

            // Configure configuration provider
            services.AddSingleton<ICredentialProvider, ConfigurationCredentialProvider>();

            // Create the storage we'll be using for User and Conversation state. (Memory is great for testing purposes.)
            services.AddSingleton<IStorage, MemoryStorage>();

            // Create the User state. (Used to store the user's language preference.)
            services.AddSingleton<UserState>();
            services.AddSingleton<ConversationState>();

            services.AddSingleton<PersonalChatMainDialog>();
            services.AddTransient<ChangeLanguageDialog>();

            // Create the Microsoft Translator responsible for making calls to the Cognitive Services translation service
            services.AddSingleton<TranslatorService>();
            services.AddSingleton<TranslationSettings>();

            // Create the Translation Middleware that will be added to the middleware pipeline in the AdapterWithErrorHandler
            services.AddSingleton<TranslationMiddleware>();

            services.AddSingleton<ICredentialProvider, ConfigurationCredentialProvider>();
            services.AddSingleton<ITicketsProvider>(new TicketsProvider(this.Configuration["StorageConnectionString"]));
            services.AddSingleton<IBatchFileProvider>(svc =>
                {
                    var telemetryClient = svc.GetService<TelemetryClient>();
                    return new BlobBatchFileStorageProvider(this.Configuration["StorageConnectionString"], telemetryClient);
                });
            // services.AddSingleton<IBotFrameworkHttpAdapter, BotFrameworkHttpAdapter>();
            services.AddSingleton<IBotFrameworkHttpAdapter, AdapterWithErrorHandler>();
            services.AddSingleton(new MicrosoftAppCredentials(this.Configuration["MicrosoftAppId"], this.Configuration["MicrosoftAppPassword"]));

            IQnAMakerClient qnaMakerClient = new QnAMakerClient(
                new ApiKeyServiceClientCredentials(this.Configuration["QnAMakerSubscriptionKey"]))
            { Endpoint = StripRouteFromQnAMakerEndpoint(this.Configuration["QnAMakerApiEndpointUrl"]) };
            string endpointKey = this.Configuration["PrimaryEndpointKey"]; //Task.Run(() => qnaMakerClient.EndpointKeys.GetKeysAsync()).Result.PrimaryEndpointKey;

            services.AddSingleton<IQnaServiceProvider>((provider) => new QuestionAnsweringQnaProvider(
                provider.GetRequiredService<IConfiguration>(),
                provider.GetRequiredService<Common.Providers.IConfigurationDataProvider>(),
                provider.GetRequiredService<IOptionsMonitor<QnAMakerSettings>>()));
            services.AddSingleton<IActivityStorageProvider>((provider) => new ActivityStorageProvider(provider.GetRequiredService<IOptionsMonitor<KnowledgeBaseSettings>>()));
            services.AddSingleton<IKnowledgeBaseSearchService>((provider) => new KnowledgeBaseSearchService(this.Configuration["SearchServiceName"], this.Configuration["SearchServiceQueryApiKey"], this.Configuration["SearchServiceAdminApiKey"], this.Configuration["StorageConnectionString"]));
            services.AddSingleton<IImageStorageProvider>((provider) => new ImageStorageProvider(this.Configuration["StorageConnectionString"]));

            services.AddSingleton<ISearchService, SearchService>();
            services.AddSingleton<IMemoryCache, MemoryCache>();
            services.AddTransient(sp => (BotFrameworkAdapter)sp.GetRequiredService<IBotFrameworkHttpAdapter>());
            services.AddTransient<IBot, FaqPlusPlusBot>();

            // Create the telemetry middleware(used by the telemetry initializer) to track conversation events
            services.AddSingleton<TelemetryLoggerMiddleware>();
            services.AddMemoryCache();
        }
    }
}