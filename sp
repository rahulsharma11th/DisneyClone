using Microsoft.SharePoint.Client;
using Microsoft.Identity.Client;
using System;
using System.Threading.Tasks;

namespace SharePointUserProfileRetrieval
{
    class Program
    {
        static async Task Main(string[] args)
        {
            // Azure AD app registration details and SharePoint site URL
            string clientId = "your-client-id";
            string tenantId = "your-tenant-id";
            string clientSecret = "your-client-secret";
            string siteUrl = "https://yourtenant.sharepoint.com";
            string userEmail = "user-email@yourtenant.com"; // User email to fetch profile

            // Get access token
            var accessToken = await GetAccessToken(clientId, tenantId, clientSecret);

            // Fetch user profile
            await FetchUserProfile(siteUrl, accessToken, userEmail);
        }

        static async Task<string> GetAccessToken(string clientId, string tenantId, string clientSecret)
        {
            IConfidentialClientApplication app = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithClientSecret(clientSecret)
                .WithAuthority(new Uri($"https://login.microsoftonline.com/{tenantId}/v2.0"))
                .Build();

            var scopes = new string[] { "https://graph.microsoft.com/.default" };
            AuthenticationResult result = await app.AcquireTokenForClient(scopes).ExecuteAsync();
            return result.AccessToken;
        }

        static async Task FetchUserProfile(string siteUrl, string accessToken, string userEmail)
        {
            using (var context = new ClientContext(siteUrl))
            {
                context.ExecutingWebRequest += (sender, e) =>
                {
                    e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + accessToken;
                };

                PeopleManager peopleManager = new PeopleManager(context);
                PersonProperties personProperties = peopleManager.GetPropertiesFor(userEmail);

                context.Load(personProperties);
                await context.ExecuteQueryAsync();

                // Here you can access various user profile properties
                Console.WriteLine("User Display Name: " + personProperties.DisplayName);
                // Add more properties as needed
            }
        }
    }
}
