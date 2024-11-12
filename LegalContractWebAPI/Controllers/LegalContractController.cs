using LegalContractWebAPI.Common;
using LegalContractWebAPI.Models;
using Microsoft.AspNetCore.Mvc;
using System.IO.Compression;
using System.Net;
using System.Text;
using System.Text.Json;
using OpenAI.Assistants;
using DocumentFormat.OpenXml.Packaging;

namespace LegalContractWebAPI.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class LegalContractController : ControllerBase
    {
        private readonly ILogger<LegalContractController> _logger;
        private readonly IConfiguration _configuration;

        public LegalContractController(ILogger<LegalContractController> logger, IConfiguration configuration)
        {
            _logger = logger;
            _configuration = configuration;
        }

        #region Controller methods

        // POST api/files
        [HttpPost("ProcessArticle")]
        public async Task<IActionResult> ProcessArticle([FromBody] ProcessDTO data)
        {

            await Task.Delay(1000); // Add 3 second delay
            if (data.scope == "document")
            {
                var response2 = new
                {
                    message = "Article processed successfully",
                    htmlString = "<html>\r\n<h1>5. Udemy’s Rights to Content You Post</h1>\r\n<p>You retain ownership of content you post to our platform, including your courses. Udemy is <span style=\"background-color: yellow;\">not allowed to share your content without your explicit consent to anyone through any media, including</span> promoting it via advertising on other websites.</p>\r\n<p>The content you post as a student or instructor (including courses) remains yours. By posting courses and other content, you do not allow Udemy to <span style=\"background-color: yellow;\">reuse and share it and you do not lose any ownership rights you may have over your content.</span> If you are an instructor, be sure to understand the content licensing terms that are detailed in the Instructor Terms.</p>\r\n<p>When you post content, comments, questions, reviews, and when you submit to us ideas and suggestions for new features or improvements, <span style=\"background-color: yellow;\">you do not authorize Udemy to use and share this content with anyone, distribute it and promote it on any platform and in any media, and you retain the right to control all modifications or edits made to it.</span></p>\r\n<p>In legal language, by submitting or posting content on or through the platforms, you grant us <span style=\"background-color: yellow;\">no license to use, copy, reproduce, process, adapt, modify, publish, transmit, display, and distribute your content</span> (including your name and image) in any and all media or distribution methods (existing now or later developed).</p>\r\n\r\n<h1>8.4 Payments and Billing</h1>\r\n<p>The subscription fee will be listed at the time of your purchase. You can visit our Support Page to learn more about where to find the fees and dates applicable to your subscription. We may also be required to add taxes to your subscription fee as described in the “Payments, Credits, and Refunds” section above. Payments are non-refundable and there are no refunds or credits for partially used periods, unless otherwise required by applicable law. <span style=\"background-color: lightgreen;\">However, users from the EU have the right to request a refund at any time within a 14-day period.</span> Depending on where you are located, you may qualify for a refund. See our Refund Policy for Subscription Plans for additional information.</p>\r\n<p>To subscribe to a Subscription Plan, you must provide a payment method. By subscribing to a Subscription Plan and providing your billing information during checkout, you grant us and our payment service providers the right to process payment for the then-applicable fees via the payment method we have on record for you. At the end of each subscription term, we will automatically renew your subscription for the same length of term and process your payment method for payment of the then-applicable fees.</p>\r\n<p>In the event that we update your payment method using information provided by our payment service providers (as described in the “Payments, Credits, and Refunds” section above), you authorize us to continue to charge the then-applicable fees to your updated payment method.</p>\r\n<p>If we are unable to process payment through the payment method we have on file for you, or if you file a chargeback disputing charges made to your payment method and the chargeback is granted, we may suspend or terminate your subscription.</p>\r\n<p>The <span style=\"background-color: lightgreen;\">subscription plan can only be changed by the user.</span></p>\r\n\r\n<h1>9.3 Limitation of Liability</h1>\r\n<p>There are risks inherent to using our Services, for example, if you access health and wellness content like yoga, and you injure yourself. You fully accept these risks and you agree that you will have no recourse to seek damages against even if you suffer loss or damage from using our platform and Services. In legal, more complete language, to the extent permitted by law, we (and our group companies, suppliers, partners, and agents) <span style=\"background-color: yellow;\">will be liable for any indirect, incidental, punitive, or consequential damages</span> (including loss of data, revenue, profits, or business opportunities, or personal injury or death), <span style=\"background-color: yellow;\">without any limitation on liability</span> and even if we’ve been advised of the possibility of damages in advance. Our liability (and the liability of each of our group companies, suppliers, partners, and agents) to you or any third parties under any circumstance shall <span style=\"background-color: yellow;\">not be limited to the greater of $100 USD or the amount you have paid us in the 12 months before the event giving rise to your claims.</span> Some jurisdictions don’t allow the exclusion or limitation of liability for consequential or incidental damages, so some of the above may not apply to you.</p>\r\n</html>" +
                    "</body></html>",
                    assistantId = "1111111111111111111111"
                };

                return new JsonResult(response2);
            }
            else
            {
                var response = new
                {
                    message = "Article processed successfully",
                    htmlString = "<html><body><p>We reserve the right to change our Subscription Plans or adjust pricing for our Services at our sole discretion. <span style='background-color: yellow;'>Any changes to your subscription plan can only be made by the user.</span> Any price changes or changes to your subscription will take effect following notice to you, except as otherwise required by applicable law.</p>" +
                      "</body></html>",
                    assistantId = "1111111111111111111111"
                };

                return new JsonResult(response);
            }
            try
            {
        // Process the article and annotations here
                Assistant assistant = null;
        var apiKey = _configuration["APIKey"];
        var helper = new OpenAIHelper(_logger, _configuration);
        if (String.IsNullOrEmpty(data.assistantId)) //new session - create new assistant
        {
            string contractFilePath = string.Empty;

            //save the file content to a temp file
                    if (!String.IsNullOrEmpty(data.fileContent))
            {
                contractFilePath = await CreateFileFromContent(data.fileContent);
                if (!String.IsNullOrEmpty(apiKey))
                {
                    assistant = await helper.CreateNewAssistant(contractFilePath, apiKey);
                }
            }
        }
        else
        {
            //use existing assistant
                    if (!String.IsNullOrEmpty(apiKey))
            {
                assistant = await helper.GetAssistant(data.assistantId, apiKey);
            }
        }
        var changeInstructions = string.Empty;
        var grouped = data.articles.GroupBy(x => x.article).ToDictionary(x => x.Key, x => x.SelectMany(y => y.annotations).ToList());
        foreach (var art in grouped)
        {
            changeInstructions += $"Change Article '{art.Key}' in a way that {string.Join(" and also ", art.Value)}. {Environment.NewLine}";
        }
        var question = string.Empty;
        if (data.scope == "document")
        {
            question = changeInstructions +
            "Show only revised version. Convert text to HTML. Use <html> tag for beginning of the text. " +
            "Use </html> tag for the end of the text.Use <p> and </p> tags for beginning and ending of each paragraph. " +
            "Include article header into response. Use <h1> and </h1> tags for beginning and ending of each header." +
            "Highlight changed text with yellow background. " +
            "Highlight added text with light green background. Do not show article name. Convert original links to html.";
        }
        else
        {
            question = changeInstructions +
            "Show only revised version of this article. Convert text to HTML. Use <html> tag for beginning of the text. " +
            "Use </html> tag for the end of the text.Use <p> and </p> tags for beginning and ending of each paragraph. " +
            "Highlight changed text with yellow background. " +
            "Highlight added text with light green background. Do not show article name. Convert original links to html.";
        }
        if (assistant == null)
        {
            var errorResponse = new
            {
                message = "An error occurred while creating assistant",
                details = "Assistant is null"
            };
            _logger.LogError("An error occurred while creating assistant. Assistant is null");
            return StatusCode((int)HttpStatusCode.BadRequest, JsonSerializer.Serialize(errorResponse));
        }
        var openAIresponse = await helper.GenerateResponse(question, assistant.Id, apiKey);


        //// Generate the HTML string
        //string htmlString = await Task.FromResult("<html><body><h1>Processed Article</h1><p>" + data.article + "</p></body></html>");

                var response = new
                {
                    message = "Article processed successfully",
                    htmlString = openAIresponse,
                    assistantId = assistant.Id
                };
        _logger.LogInformation($"Article processed successfully. Session = {assistant.Id}");
        return new JsonResult(response);
    }
            catch (Exception ex)
            {
                var errorResponse = new
                {
                    message = "An error occurred while processing the article",
                    details = ex.Message
                };
    _logger.LogError(ex, "An error occurred while processing the article");
                return StatusCode((int)HttpStatusCode.BadRequest, JsonSerializer.Serialize(errorResponse));
}
        }





        // POST 
        [HttpPost("DeleteAssistant")]
public async Task<IActionResult> DeleteAssistant([FromBody] string assistantId)
{
    try
    {
        var apiKey = _configuration["APIKey"];
        var helper = new OpenAIHelper(_logger, _configuration);
        var response = await helper.DeleteAssistant(assistantId, apiKey);
        _logger.LogInformation($"Assistant {assistantId} deleted successfully");
        return new JsonResult(response);
    }
    catch (Exception ex)
    {
        var errorResponse = new
        {
            message = "An error occurred while deleting assistant",
            details = ex.Message
        };
        _logger.LogError(ex, $"An error occurred while deleting Assistant {assistantId}");
        return StatusCode((int)HttpStatusCode.BadRequest, JsonSerializer.Serialize(errorResponse));
    }
}




#endregion
#region Helper Methods
        private async Task<String> CreateFileFromContent(string content)
{
    try
    {
        var openXML = string.Empty;
        byte[] base64String = Convert.FromBase64String(content);
        Stream docxStream = new MemoryStream(base64String);
        ZipArchive archive = new ZipArchive(docxStream);
        foreach (ZipArchiveEntry entry in archive.Entries)
        {
            // The text content of a docx is stored in word/document.xml. So we seek the data from there.
            // You may do further operation with OpenXmlApi.
                    if (entry.FullName == "word/document.xml")
            {
                Stream documentXmlData = entry.Open();

                // Convert documentXmlData to docx
                        string docxFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");
                using (WordprocessingDocument doc = WordprocessingDocument.Create(docxFilePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = doc.AddMainDocumentPart();
                    using (StreamWriter streamWriter = new StreamWriter(mainPart.GetStream()))
                    {
                        await documentXmlData.CopyToAsync(streamWriter.BaseStream);
                    }
                }
                var plainText = ReadWordDocument(docxFilePath);
                System.IO.File.Delete(docxFilePath);

                // Save as plain text file

                        string txtFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".txt");
                using (StreamWriter writer = new StreamWriter(txtFilePath))
                {
                    await writer.WriteAsync(plainText);
                }
                return await Task.FromResult(txtFilePath);
            }
        }
        return null;
    }
    catch (Exception ex)
    {
        _logger.LogError(ex, "An error occurred while creating file from content");
        throw ex;
    }
}
private string GetPlainText(DocumentFormat.OpenXml.OpenXmlElement element)
{
    StringBuilder plainTextInWord = new StringBuilder();
    foreach (DocumentFormat.OpenXml.OpenXmlElement section in element.Elements())
    {
        switch (section.LocalName)
        {
            // Text 
                    case "t":
                plainTextInWord.Append(section.InnerText);
                break;

            // Carriage return 
                    case "cr":
            case "br":
                // Page break 
                        plainTextInWord.Append(Environment.NewLine);
                        break;

                    // Tab 
                    case "tab":
                        plainTextInWord.Append("\t");
                        break;

                    // Paragraph 
                    case "p":
                        plainTextInWord.Append(GetPlainText(section));
                        plainTextInWord.AppendLine(Environment.NewLine);
                        break;

                    default:
                        plainTextInWord.Append(GetPlainText(section));
                        break;
                }
            }

            return plainTextInWord.ToString();
        }

        private string ReadWordDocument(string filePath)
        {
            StringBuilder sb = new StringBuilder();
            // Open a WordprocessingDocument for editing using the filepath.
            using (WordprocessingDocument wpd = WordprocessingDocument.Open(filePath, true))
            {
                DocumentFormat.OpenXml.OpenXmlElement element = wpd.MainDocumentPart.Document.Body;
                if (element != null)
                {
                    sb.Append(GetPlainText(element));
                }
            }

            return sb.ToString();
        }
        #endregion

        #region Tests
        // GET api/files
        [HttpGet("GetAPIKey")]
        public async Task<IActionResult> GetAPIKey()
        {
            string result = String.Empty;
            try
            {
                var apiKey = _configuration["APIKey"];
                result = await Task.FromResult(apiKey);
                return new ObjectResult(result);
            }
            catch (Exception ex)
            {
                var errorResponse = new
                {
                    message = "An error occurred while getting API key",
                    details = ex.Message
                };
                return StatusCode((int)HttpStatusCode.BadRequest, JsonSerializer.Serialize(errorResponse));

            }

        }




        //// POST api/files
        //[HttpPost("PostFileContent")]
        //public async Task<IActionResult> PostFileContent([FromBody] string content)
        //{
        //    string result = String.Empty;
        //    try
        //    {
        //        var openXML = string.Empty;

        //        byte[] base64String = Convert.FromBase64String(content);
        //        Stream docxStream = new MemoryStream(base64String);
        //        ZipArchive archive = new ZipArchive(docxStream);

        //        foreach (ZipArchiveEntry entry in archive.Entries)
        //        {
        //            // The text content of a docx is stored in word/document.xml. So we seek the data from there.
        //            // You may do further operation with OpenXmlApi.
        //            if (entry.FullName == "word/document.xml")
        //            {
        //                Stream documentXmlData = entry.Open();

        //                // Convert documentXmlData to docx
        //                string docxFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");

        //                using (WordprocessingDocument doc = WordprocessingDocument.Create(docxFilePath, WordprocessingDocumentType.Document))
        //                {
        //                    MainDocumentPart mainPart = doc.AddMainDocumentPart();

        //                    using (StreamWriter streamWriter = new StreamWriter(mainPart.GetStream()))
        //                    {
        //                        await documentXmlData.CopyToAsync(streamWriter.BaseStream);
        //                    }
        //                }
        //                var plainText = ReadWordDocument(docxFilePath);

        //                System.IO.File.Delete(docxFilePath);

        //                // Save as plain text file

        //                string txtFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".txt");

        //                using (StreamWriter writer = new StreamWriter(txtFilePath))
        //                {
        //                    await writer.WriteAsync(plainText);
        //                }

        //                //



        //                System.IO.File.Delete(txtFilePath);


        //                break;
        //            }
        //        }

        //        result = await Task.FromResult("String posted successfully");
        //        return new ObjectResult(result);
        //    }
        //    catch (Exception ex)
        //    {
        //        var errorResponse = new
        //        {
        //            message = "An error occurred while processing the string",
        //            details = ex.Message
        //        };
        //        return StatusCode((int)HttpStatusCode.BadRequest, JsonSerializer.Serialize(errorResponse));
        //    }
        //}
        #endregion

    }
}
