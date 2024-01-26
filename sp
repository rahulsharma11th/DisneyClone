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

/////////////////


using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SharePointProfileRetrieval
{
    class Program
    {
        static async Task Main(string[] args)
        {
            string siteUrl = "https://yourtenant.sharepoint.com";
            string userName = "user@yourtenant.com"; // Main user's email address

            using (var context = new ClientContext(siteUrl))
            {
                // Assuming context is already authenticated
                await FetchUserProfileAndReports(context, userName, 0);
            }
        }

        static async Task FetchUserProfileAndReports(ClientContext context, string accountName, int level)
        {
            PeopleManager peopleManager = new PeopleManager(context);
            PersonProperties personProperties = peopleManager.GetPropertiesFor(accountName);

            context.Load(personProperties, p => p.DirectReports, p => p.DisplayName, p => p.Email, p => p.Title, p => p.UserProfileProperties);
            await context.ExecuteQueryAsync();

            // Print user profile
            string indent = new String(' ', level * 2);
            Console.WriteLine($"{indent}Name: {personProperties.DisplayName}");
            Console.WriteLine($"{indent}Email: {personProperties.Email}");
            Console.WriteLine($"{indent}Job Title: {personProperties.Title}");
            // Add more properties as needed

            // Fetch and print direct reports
            foreach (var report in personProperties.DirectReports)
            {
                // Assuming report.Email contains the account name
                await FetchUserProfileAndReports(context, report.Email, level + 1);
            }
        }
    }
}

