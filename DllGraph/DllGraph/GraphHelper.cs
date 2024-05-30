using Azure.Identity;
using Microsoft.Graph.Models.ODataErrors;
using Microsoft.Graph.Models;
using Microsoft.Graph;
using System.Runtime.CompilerServices;


namespace DllGraph
{
    public class GraphHelper
    {
        public static async Task<MessageCollectionResponse?> GetInboxAsync(string TenantId, string ClientId, string ClientSecret, string EmailAdress)
        {
                var scopes = new[] { "https://graph.microsoft.com/.default" };

                var clientSecretCredential = new ClientSecretCredential(
                    TenantId, ClientId, ClientSecret);

                var graphClient = new GraphServiceClient(clientSecretCredential, scopes);


                var result = await graphClient.Users[EmailAdress].MailFolders["Inbox"].Messages.GetAsync((config) =>
                {
                    // Only request specific properties
                    config.QueryParameters.Select = new[] { "from", "receivedDateTime", "subject", "sender" };

                    // Get at most 25 results
                    config.QueryParameters.Top = 25;
                    // Sort by received time, newest first
                    config.QueryParameters.Orderby = new[] { "receivedDateTime DESC" };
                });

                return result;

        }

        public static async Task<MessageCollectionResponse?> GetAttachmentsAsync(string TenantId, string ClientId, string ClientSecret, string EmailAdress)
        {

                var scopes = new[] { "https://graph.microsoft.com/.default" };

                var clientSecretCredential = new ClientSecretCredential(
                    TenantId, ClientId, ClientSecret);

                var graphClient = new GraphServiceClient(clientSecretCredential, scopes);


                var result = await graphClient.Users[EmailAdress].MailFolders["Inbox"].Messages.GetAsync((config) =>
                {
                    // Only request specific properties
                    config.QueryParameters.Select = new[] { "HasAttachments" };

                    config.QueryParameters.Expand = new[] { "Attachments" };

                });

                return result;
           

        }

        public static async Task DeleteInboxByIdAsync(string idMessage, string TenantId, string ClientId, string ClientSecret, string EmailAdress)
        {
                var scopes = new[] { "https://graph.microsoft.com/.default" };

                var clientSecretCredential = new ClientSecretCredential(
                    TenantId, ClientId, ClientSecret);

                var graphClient = new GraphServiceClient(clientSecretCredential, scopes);



                await graphClient.Users[EmailAdress].MailFolders["Inbox"].Messages[idMessage].DeleteAsync();

                // return await result.;        

        }


    }
}