// <copyright file="QnaServiceProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.FAQPlusPlus.Common.Providers
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Web;
    using Microsoft.Azure.CognitiveServices.Knowledge.QnAMaker;
    using Microsoft.Azure.CognitiveServices.Knowledge.QnAMaker.Models;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.FAQPlusPlus.Common.Models;
    using Microsoft.Teams.Apps.FAQPlusPlus.Common.Models.Configuration;

    /// <summary>
    /// Qna maker service provider class.
    /// </summary>
    public partial class QnaServiceProvider : IQnaServiceProvider
    {
        /// <summary>
        /// Environment type.
        /// </summary>
        private const string TestEnvironmentType = "Prod";
        private const string ProdEnvironmentType = "Test";

        private readonly IConfigurationDataProvider configurationProvider;
        private readonly IQnAMakerClient qnaMakerClient;
        private readonly QnAMakerRuntimeClient qnaMakerRuntimeClient;

        /// <summary>
        /// Represents a set of key/value application configuration properties.
        /// </summary>
        private readonly QnAMakerSettings options;

        /// <summary>
        /// Initializes a new instance of the <see cref="QnaServiceProvider"/> class.
        /// </summary>
        /// <param name="configurationProvider">ConfigurationProvider fetch and store information in storage table.</param>
        /// <param name="optionsAccessor">A set of key/value application configuration properties.</param>
        /// <param name="qnaMakerClient">Qna service client.</param>
        /// <param name="qnaMakerRuntimeClient">Qna service runtime client.</param>
        public QnaServiceProvider(IConfigurationDataProvider configurationProvider, IOptionsMonitor<QnAMakerSettings> optionsAccessor, IQnAMakerClient qnaMakerClient, QnAMakerRuntimeClient qnaMakerRuntimeClient)
        {
            this.configurationProvider = configurationProvider;
            this.qnaMakerClient = qnaMakerClient;
            this.options = optionsAccessor.CurrentValue;
            this.qnaMakerRuntimeClient = qnaMakerRuntimeClient;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="QnaServiceProvider"/> class.
        /// </summary>
        /// <param name="configurationProvider">ConfigurationProvider fetch and store information in storage table.</param>
        /// <param name="optionsAccessor">A set of key/value application configuration properties.</param>
        /// <param name="qnaMakerClient">Qna service client.</param>
        public QnaServiceProvider(IConfigurationDataProvider configurationProvider, IOptionsMonitor<QnAMakerSettings> optionsAccessor, IQnAMakerClient qnaMakerClient)
        {
            this.configurationProvider = configurationProvider;
            this.qnaMakerClient = qnaMakerClient;
            this.options = optionsAccessor.CurrentValue;
        }

        /// <summary>
        /// This method is used to add QnA pair in Kb.
        /// </summary>
        /// <param name="question">Question text.</param>
        /// <param name="combinedDescription">Answer text.</param>
        /// <param name="createdBy">Created by user.</param>
        /// <param name="conversationId">Conversation id.</param>
        /// <param name="activityReferenceId">Activity reference id refer to activityid in storage table.</param>
        /// <returns>Operation state as task.</returns>
        public async Task<Operation> AddQnaAsync(string question, string combinedDescription, string createdBy, string conversationId, string activityReferenceId)
        {
            var knowledgeBase = await this.configurationProvider.GetSavedEntityDetailAsync(ConfigurationEntityTypes.KnowledgeBaseId).ConfigureAwait(false);

            // Update knowledgebase.
            return await this.qnaMakerClient.Knowledgebase.UpdateAsync(knowledgeBase, new UpdateKbOperationDTO
            {
                // Create JSON of changes.
                Add = new UpdateKbOperationDTOAdd
                {
                    QnaList = new List<QnADTO>
                    {
                         new QnADTO
                         {
                            Questions = new List<string> { question?.Trim() },
                            Answer = combinedDescription?.Trim(),
                            Metadata = new List<MetadataDTO>()
                            {
                                new MetadataDTO() { Name = Constants.MetadataCreatedAt, Value = DateTime.UtcNow.Ticks.ToString(CultureInfo.InvariantCulture) },
                                new MetadataDTO() { Name = Constants.MetadataCreatedBy, Value = createdBy },
                                new MetadataDTO() { Name = Constants.MetadataConversationId, Value = HttpUtility.UrlEncode(conversationId) },
                                new MetadataDTO() { Name = Constants.MetadataActivityReferenceId, Value = activityReferenceId },
                            },
                         },
                    },
                },
                Update = null,
                Delete = null,
            }).ConfigureAwait(false);
        }

        /// <summary>
        /// Update Qna pair in knowledge base.
        /// </summary>
        /// <param name="questionId">Question id.</param>
        /// <param name="answer">Answer text.</param>
        /// <param name="updatedBy">Updated by user.</param>
        /// <param name="updatedQuestion">Updated question text.</param>
        /// <param name="question">Original question text.</param>
        /// <returns>Perfomed action task.</returns>
        public async Task UpdateQnaAsync(int questionId, string answer, string updatedBy, string updatedQuestion, string question, string conversationId = null, string activityReferenceId = null)
        {
            var knowledgeBaseId = await this.configurationProvider.GetSavedEntityDetailAsync(ConfigurationEntityTypes.KnowledgeBaseId).ConfigureAwait(false);
            var questions = default(UpdateQnaDTOQuestions);
            if (!string.IsNullOrEmpty(updatedQuestion?.Trim()))
            {
                questions = (updatedQuestion?.ToUpperInvariant().Trim() == question?.ToUpperInvariant().Trim()) ? null
                    : new UpdateQnaDTOQuestions()
                    {
                        Add = new List<string> { updatedQuestion.Trim() },
                        Delete = new List<string> { question.Trim() },
                    };
            }

            // Update knowledgebase.
            List<MetadataDTO> metadataDTOs = new List<MetadataDTO>()
                                {
                                    new MetadataDTO() { Name = Constants.MetadataUpdatedAt, Value = DateTime.UtcNow.Ticks.ToString(CultureInfo.InvariantCulture) },
                                    new MetadataDTO() { Name = Constants.MetadataUpdatedBy, Value = updatedBy },
                                };

            if (activityReferenceId != null)
            {
                metadataDTOs.Add(new MetadataDTO() { Name = Constants.MetadataActivityReferenceId, Value = activityReferenceId });
            }

            if (conversationId != null)
            {
                metadataDTOs.Add(new MetadataDTO() { Name = Constants.MetadataConversationId, Value = HttpUtility.UrlEncode(conversationId) });
                metadataDTOs.Add(new MetadataDTO() { Name = Constants.MetadataCreatedAt, Value = DateTime.UtcNow.Ticks.ToString(CultureInfo.InvariantCulture) });
            }

            await qnaMakerClient.Knowledgebase.UpdateAsync(knowledgeBaseId, new UpdateKbOperationDTO
            {
                // Create JSON of changes.
                Add = null,
                Update = new UpdateKbOperationDTOUpdate()
                {
                    QnaList = new List<UpdateQnaDTO>()
                    {
                        new UpdateQnaDTO()
                        {
                            Id = questionId,
                            Source = Constants.Source,
                            Answer = answer?.Trim(),
                            Questions = questions,
                            Metadata = new UpdateQnaDTOMetadata()
                            {
                                Add = metadataDTOs,
                            },
                        },
                    },
                },
                Delete = null,
            }).ConfigureAwait(false);
        }

        /// <summary>
        /// This method is used to delete Qna pair from KB.
        /// </summary>
        /// <param name="questionId">Question id.</param>
        /// <returns>Perfomed action task.</returns>
        public async Task DeleteQnaAsync(int questionId)
        {
            var knowledgeBaseId = await this.configurationProvider.GetSavedEntityDetailAsync(ConfigurationEntityTypes.KnowledgeBaseId).ConfigureAwait(false);

            // to delete a question and answer based on id.
            await this.qnaMakerClient.Knowledgebase.UpdateAsync(knowledgeBaseId, new UpdateKbOperationDTO
            {
                // Create JSON of changes.
                Add = null,
                Update = null,
                Delete = new UpdateKbOperationDTODelete()
                {
                    Ids = new List<int?>() { questionId },
                },
            }).ConfigureAwait(false);
        }

        /// <summary>
        /// Get answer from knowledgebase for a given question.
        /// </summary>
        /// <param name="question">Question text.</param>
        /// <param name="isTestKnowledgeBase">Prod or test.</param>
        /// <param name="previousQnAId">Id of previous question.</param>
        /// <param name="previousUserQuery">Previous question information.</param>
        /// <returns>QnaSearchResultList result as response.</returns>
        public async Task<QnASearchResultList> GenerateAnswerAsync(string question, bool isTestKnowledgeBase, string previousQnAId = null, string previousUserQuery = null, IList<QueryTag> tags = null)
        {
            var knowledgeBaseId = await this.configurationProvider.GetSavedEntityDetailAsync(ConfigurationEntityTypes.KnowledgeBaseId).ConfigureAwait(false);

            QueryDTO queryDTO = new QueryDTO()
            {
                IsTest = isTestKnowledgeBase,
                Question = question?.Trim(),
                ScoreThreshold = Convert.ToDouble(this.options.ScoreThreshold, CultureInfo.InvariantCulture),
                StrictFilters = tags?.Select(x => new MetadataDTO(x.Name, x.Value)).ToList(),
            };

            if (previousQnAId != null && previousUserQuery != null)
            {
                queryDTO.Context = new QueryDTOContext
                {
                    PreviousQnaId = Convert.ToInt32(previousQnAId),
                    PreviousUserQuery = previousUserQuery,
                };
            }

            QnASearchResultList qnASearchResultList = await this.qnaMakerClient.Knowledgebase.GenerateAnswerAsync(knowledgeBaseId, queryDTO).ConfigureAwait(false);
            return qnASearchResultList;
        }

        /// <summary>
        /// This method returns the downloaded knowledgebase documents.
        /// </summary>
        /// <param name="knowledgeBaseId">Knowledgebase Id.</param>
        /// <returns>List of question and answer document object.</returns>
        public async Task<IEnumerable<QnADTO>> DownloadKnowledgebaseAsync(string knowledgeBaseId, bool isTestEnvironment = false)
        {
            var qnaDocuments = await this.qnaMakerClient.Knowledgebase.DownloadAsync(
                knowledgeBaseId,
                environment: isTestEnvironment ? TestEnvironmentType : ProdEnvironmentType
                ).ConfigureAwait(false);
            return qnaDocuments.QnaDocuments;
        }

        /// <summary>
        /// Checks whether knowledgebase need to be published.
        /// </summary>
        /// <param name="knowledgeBaseId">Knowledgebase id.</param>
        /// <returns>A <see cref="Task"/> of type bool where true represents knowledgebase need to be published while false indicates knowledgebase not need to be published.</returns>
        public async Task<bool> GetPublishStatusAsync(string knowledgeBaseId)
        {
            var qnaDocuments = await this.qnaMakerClient.Knowledgebase.GetDetailsAsync(knowledgeBaseId).ConfigureAwait(false);
            if (qnaDocuments != null && qnaDocuments.LastChangedTimestamp != null && qnaDocuments.LastPublishedTimestamp != null)
            {
                return Convert.ToDateTime(qnaDocuments.LastChangedTimestamp) > Convert.ToDateTime(qnaDocuments.LastPublishedTimestamp);
            }

            return true;
        }

        /// <summary>
        /// Method is used to publish knowledgebase.
        /// </summary>
        /// <param name="knowledgeBaseId">Knowledgebase Id.</param>
        /// <returns>Task for published data.</returns>
        public async Task PublishKnowledgebaseAsync(string knowledgeBaseId)
        {
            await this.qnaMakerClient.Knowledgebase.PublishAsync(knowledgeBaseId).ConfigureAwait(false);
        }

        /// <summary>
        /// Get knowledgebase published information.
        /// </summary>
        /// <param name="knowledgeBaseId">Knowledgebase id.</param>
        /// <returns>A <see cref="Task"/> of type bool where true represents knowledgebase has published atleast once while false indicates that knowledgebase has not published yet.</returns>
        public async Task<bool> GetInitialPublishedStatusAsync(string knowledgeBaseId)
        {
            KnowledgebaseDTO qnaDocuments = await this.qnaMakerClient.Knowledgebase.GetDetailsAsync(knowledgeBaseId).ConfigureAwait(false);
            return !string.IsNullOrEmpty(qnaDocuments.LastPublishedTimestamp);
        }
    }
}
