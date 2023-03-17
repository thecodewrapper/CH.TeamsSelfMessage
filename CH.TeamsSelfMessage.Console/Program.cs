using System;
using System.IO;
using System.Threading.Tasks;
using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace TeamsSelfMessage
{
    class Program
    {
        private static string[] _graphScopes = new[] { "User.Read", "ChatMessage.Send", "Chat.ReadWrite" };
        private const string SELF_CHAT_ID = "48:notes";

        public static async Task Main(string[] args)
        {
            var configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();

            var graphClient = await GetGraphClient(configuration);

            string messageContent = "This is a message to myself!";
            ChatMessage sentMessage = await SendMessageAsync(graphClient, messageContent);
            Console.WriteLine($"Message sent with ID: {sentMessage.Id}");

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        private static async Task<GraphServiceClient> GetGraphClient(IConfiguration configuration)
        {
            var interactiveBrowserCredentialOptions = new InteractiveBrowserCredentialOptions
            {
                ClientId = configuration["appId"],
                TenantId = configuration["tenantId"]
            };
            var tokenCredential = new InteractiveBrowserCredential(interactiveBrowserCredentialOptions);

            var graphClient = new GraphServiceClient(tokenCredential, _graphScopes);
            _ = await graphClient.Me.GetAsync(); //trigger login
            return graphClient;
        }

        private static async Task<ChatMessage> SendMessageAsync(GraphServiceClient graphClient, string messageContent)
        {
            var message = new ChatMessage
            {
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = messageContent
                }
            };

            return await graphClient.Me.Chats[SELF_CHAT_ID].Messages.PostAsync(message);
        }
    }
}