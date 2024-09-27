using LegalContractWebAPI.Controllers;
using Microsoft.AspNetCore.DataProtection.KeyManagement;
using OpenAI;
using OpenAI.Assistants;
using OpenAI.Files;
using System.ClientModel;
using System.Text;


namespace LegalContractWebAPI.Common
{
    public class OpenAIHelper
    {
        private readonly ILogger<LegalContractController> _logger;
        private readonly IConfiguration _configuration;
        public OpenAIHelper(ILogger<LegalContractController> logger, IConfiguration configuration)
        {
            _logger = logger;
            _configuration = configuration;
        }

        public async Task<Assistant> CreateNewAssistant(string filePath, string apiKey)
        {
            try
            {
                OpenAIClient openAIClient = new(apiKey);

                OpenAIFileClient fileClient = openAIClient.GetOpenAIFileClient();
                AssistantClient assistantClient = openAIClient.GetAssistantClient();

                using Stream document = File.OpenRead(filePath);

                var assistantName = "Assistant_" + DateTime.Now.ToString();
                //var contractFileName = "Contract_" + DateTime.Now.ToString().Replace(" ", "") + ".txt";
                var contractFileName = Path.GetFileName(filePath);

                OpenAIFile contractFile = await fileClient.UploadFileAsync(
                                                                        document,
                                                                        contractFileName,
                                                                        FileUploadPurpose.Assistants);

                // Now, we'll create a client intended to help with that data
                AssistantCreationOptions assistantOptions = new()
                {
                    Name = assistantName,
                    Instructions =
                    "You are a lawyer reviewing a contract.Change contract based on user input.Use the whole contract as context," +
                    "but only return re-written version of the specified article. Do not respond with the entire contract. " +
                    "Use file " + contractFileName + ".It is in your Files, you can retrieve it.",
                    Tools =
            {
                new FileSearchToolDefinition(),
                new CodeInterpreterToolDefinition(),
            },
                    ToolResources = new()
                    {
                        FileSearch = new()
                        {
                            NewVectorStores =
                    {
                        new VectorStoreCreationHelper([contractFile.Id]),
                    }
                        }
                    },
                };

                Assistant assistant = await assistantClient.CreateAssistantAsync("gpt-4o-mini", assistantOptions);
                return assistant;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating assistant");
                throw ex;
            }
        }

        public async Task<Assistant> GetAssistant(string assistantId, string apiKey)
        {
            try
            {

                AssistantClient client = new(apiKey);

                AsyncCollectionResult<Assistant> assistants = client.GetAssistantsAsync();

                await foreach (Assistant assistant in assistants)
                {
                    //                Console.WriteLine($"[{count,3}] {assistant.Id} {assistant.CreatedAt:s} {assistant.Name}");
                    if (assistant.Id == assistantId)
                    {
                        return assistant;
                    }
                }
                return null;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error getting assistant {assistantId}");
                throw ex;

            }
        }

        public async Task<bool> DeleteAssistant(string assistantId, string apiKey)
        {
            try
            {
                AssistantClient assistantClient = new(apiKey);

                _ = await assistantClient.DeleteAssistantAsync(assistantId);
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error deleting assistant {assistantId}");
                throw ex;

            }
        }


        public async Task<String> GenerateResponse(string question, string assistantId, string apiKey)
        {
            AssistantClient assistantClient = new(apiKey);

            // Now we'll create a thread with a user query about the data already associated with the assistant, then run it
            ThreadCreationOptions threadOptions = new()
            {
                InitialMessages = { question }
            };


            ThreadRun threadRun = await assistantClient.CreateThreadAndRunAsync(assistantId, threadOptions);

            // Check back to see when the run is done
            do
            {
                Thread.Sleep(TimeSpan.FromSeconds(1));
                threadRun = assistantClient.GetRun(threadRun.ThreadId, threadRun.Id);
            } while (!threadRun.Status.IsTerminal);

            // Finally, we'll print out the full history for the thread that includes the augmented generation
            AsyncCollectionResult<ThreadMessage> messages
                = assistantClient.GetMessagesAsync(threadRun.ThreadId, new MessageCollectionOptions() { Order = MessageCollectionOrder.Ascending });

            var result = new StringBuilder();

            var i = 0;
            await foreach (ThreadMessage message in messages)
            {

                //Console.Write($"[{message.Role.ToString().ToUpper()}]: ");
                foreach (MessageContent contentItem in message.Content)
                {
                    if (!string.IsNullOrEmpty(contentItem.Text))
                    {
                        if (i > 0) //first message is always the question - skip it
                        {
                            result.AppendLine(contentItem.Text.Replace("```html", "").Replace("```", ""));
                            if (contentItem.TextAnnotations.Count > 0)
                            {
                                result.AppendLine("");
                            }
                        }
                    }

                    // Include annotations, if any.
                    foreach (TextAnnotation annotation in contentItem.TextAnnotations)
                    {
                        if (!string.IsNullOrEmpty(annotation.InputFileId))
                        {
                            // Console.WriteLine($"* File citation, file ID: {annotation.InputFileId}");
                        }
                        if (!string.IsNullOrEmpty(annotation.OutputFileId))
                        {
                            //Console.WriteLine($"* File output, new file ID: {annotation.OutputFileId}");
                        }
                    }
                }
                i++;
            }



            //clean-up response to take html only
            var response = result.ToString();
            var htmlBegin = response.IndexOf("<html>");
            var htmlEnd = response.IndexOf("</html>");

            var htmlOnly = response.Substring(htmlBegin, htmlEnd + 7);

            return htmlOnly;
        }



    }

}
