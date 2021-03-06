﻿namespace Microsoft.Teams.Apps.FAQPlusPlus.Controllers
{
    using System;
    using System.IO;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.FAQPlusPlus.Common.Models;
    using Microsoft.Teams.Apps.FAQPlusPlus.Common.Models.Configuration;
    using Microsoft.Teams.Apps.FAQPlusPlus.Common.Providers;
    using Microsoft.Teams.Apps.FAQPlusPlus.Helpers;
    using Newtonsoft.Json;

    /// <summary>
    /// Controller to update QnA answers and upload images.
    /// </summary>
    [Route("/question")]
    public class QuestionController : Controller
    {
        private readonly IConfigurationDataProvider configurationProvider;
        private readonly IQnaServiceProvider qnaServiceProvider;
        private readonly IImageStorageProvider imageStorageProvider;
        private readonly BotSettings options;
        private readonly string appId;

        /// <summary>
        /// Initializes a new instance of the <see cref="QuestionController"/> class.
        /// </summary>
        /// <param name="configurationProvider"></param>
        /// <param name="qnaServiceProvider"></param>
        /// <param name="imageStorageProvider"></param>
        /// <param name="optionsAccessor"></param>
        public QuestionController(IConfigurationDataProvider configurationProvider, IQnaServiceProvider qnaServiceProvider, IImageStorageProvider imageStorageProvider, IOptionsMonitor<BotSettings> optionsAccessor)
        {
            this.configurationProvider = configurationProvider;
            this.qnaServiceProvider = qnaServiceProvider;
            this.imageStorageProvider = imageStorageProvider;

            this.options = optionsAccessor.CurrentValue;
            this.appId = this.options.MicrosoftAppId;
        }


        // GET: QuestionController
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// Get the QnA Pair for Editing
        /// </summary>
        /// <param name="id"></param>
        /// <param name="question"></param>
        /// <param name="answer"></param>
        /// <returns></returns>
        //// GET: QuestionController/Edit/5
        [Route("/question/edit/{id}")]
        public async Task<ActionResult> Edit(int id, string question, string answer)
        {
            var qnaModel = new QnAQuestionModel();
            AdaptiveSubmitActionData postedValues = new AdaptiveSubmitActionData();

            if (id > 0)
            {
                var knowledgeBaseId = await this.configurationProvider.GetSavedEntityDetailAsync(ConfigurationEntityTypes.KnowledgeBaseId).ConfigureAwait(false);
                var qnaitems = await this.qnaServiceProvider.DownloadKnowledgebaseAsync(knowledgeBaseId);

                var answerData = qnaitems.FirstOrDefault(k => k.Id == id);

                if (answerData != null)
                {
                    postedValues.QnaPairId = id;
                    postedValues.OriginalQuestion = answerData.Questions[0];
                    postedValues.UpdatedQuestion = answerData.Questions[0];

                    if (Validators.IsValidJSON(answerData.Answer))
                    {
                        AnswerModel answerModel = JsonConvert.DeserializeObject<AnswerModel>(answerData.Answer);
                        postedValues.Description = answerModel.Description;
                        postedValues.Title = answerModel.Title;
                        postedValues.Subtitle = answerModel.Subtitle;
                        postedValues.ImageUrl = answerModel.ImageUrl;
                        postedValues.RedirectionUrl = answerModel.RedirectionUrl;
                    }
                    else
                    {
                        postedValues.Description = answerData.Answer;
                        //postedValues.ImageUrl = "https://3uc74q2sbxzd4.blob.core.windows.net/faqplus-image-container/20210513080531_Feb1_Byron_03.jpg";
                        if (!String.IsNullOrEmpty(postedValues.ImageUrl))
                        {
                            qnaModel.ImageMd = $"![]({postedValues.ImageUrl})";
                        }
                    }
                }
                else
                {
                    postedValues.Description = "ERROR: QnA Pair Not Found";
                }
            }

            qnaModel.PostedValues = postedValues;
            qnaModel.AppId = this.appId;

            return View(qnaModel);
        }

        /// <summary>
        /// Posting edited Answer
        /// </summary>
        /// <param name="id"></param>
        /// <param name="collection"></param>
        /// <returns></returns>
        // POST: QuestionController/Edit/5
        [Route("/question/edit/{id}")]
        [HttpPost]
        public async Task<ActionResult> Edit(int id, IFormCollection collection)
        {
            try
            {
                return RedirectToAction(nameof(Index));
            }
            catch
            {
                return View();
            }
        }

        [Route("/question/image/{imageName}")]
        [HttpGet]
        public async Task<ActionResult> GetImage(string imageName)
        {
            var image = await this.imageStorageProvider.GetAsync(imageName);
            return File(image, "application/octet-stream", imageName);
        }

        /// <summary>
        /// Upload Image To Blob Storage
        /// </summary>
        /// <param name="collection"></param>
        /// <returns></returns>
        [Route("/question/upload")]
        [HttpPost]
        public async Task<ActionResult> Upload(IFormCollection collection)
        {

            string url = string.Empty;
            string fileName = string.Empty;

            Console.WriteLine(collection.Count);
            if (collection.Files.Count > 0)
            {
                if (collection.Files[0] != null)
                {
                    var file = collection.Files[0];

                    if (IsImage(file))
                    {
                        if (file.Length > 0)
                        {
                            using (Stream stream = file.OpenReadStream())
                            {
                                // Get the reference to the block blob from the container
                                string orginalFileName = file.FileName;
                                string filenamePrefix = DateTime.Now.ToString("yyyyMMddHHmmss"); // Makes filename unique
                                string newFileName = $"{filenamePrefix}_{orginalFileName}";

                                if (file.FileName.LastIndexOf("\\") > -1)
                                {
                                    orginalFileName = file.FileName.Substring(file.FileName.LastIndexOf("\\") + 1,
                                    file.FileName.Length - file.FileName.LastIndexOf("\\") - 1);
                                }

                                var storageUrl = await this.imageStorageProvider.UploadAsync(stream, newFileName);
                                url = $"{this.Request.Scheme}://{this.Request.Host}/question/image/{newFileName}";
                            }
                        }
                    }
                }
            }

            // CKEDitor requires image url to be passed in JSON
            return Json(new { Url = url });
        }



        /// <summary>
        /// Checks to see if image is one of allowable types
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        private bool IsImage(IFormFile file)
        {
            if (file.ContentType.Contains("image"))
            {
                return true;
            }

            string[] formats = new string[] { ".jpg", ".png", ".gif", ".jpeg" };

            return formats.Any(item => file.FileName.EndsWith(item, StringComparison.OrdinalIgnoreCase));
        }
    }
}
