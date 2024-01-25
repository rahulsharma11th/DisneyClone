using Microsoft.SharePoint.Client;
using Azure.Identity;
using System;
using Microsoft.SharePoint.Client.UserProfiles;
using System.Net;

class SharePointOnlineCredentials : ICredentials
{
    private ClientSecretCredential _clientSecretCredential;

    public SharePointOnlineCredentials(string clientId, string clientSecret, string tenantId)
    {
        _clientSecretCredential = new ClientSecretCredential(tenantId, clientId, clientSecret);
    }

    public System.Net.NetworkCredential GetCredential(Uri uri, string authType)
    {
        // Return null for the NetworkCredential part as it's not used.
        return null;
    }
}

class Program
{
    static async System.Threading.Tasks.Task Main(string[] args)
    {
        string siteUrl = "https://yourtenant.sharepoint.com/sites/yoursite";
        string clientId = "your-client-id";
        string clientSecret = "your-client-secret";
        string tenantId = "your-tenant-id";

        // Specify the email address for the user you want to retrieve information for
        string userEmail = "user@example.com"; // Replace with the user's email

        var spCredentials = new SharePointOnlineCredentials(clientId, clientSecret, tenantId);

        using (var context = new ClientContext(siteUrl))
        {
            context.Credentials = spCredentials;

            Web web = context.Web;
            context.Load(web);

            try
            {
                await context.ExecuteQueryAsync();

                // Get user profile properties for the specified email address
                var peopleManager = new PeopleManager(context);
                var userProfileProperties = peopleManager.GetPropertiesFor(userEmail);

                await context.ExecuteQueryAsync();

                Console.WriteLine("Web Title: " + web.Title);
                Console.WriteLine($"User Profile Data for {userEmail}:");

                // Access and print specific profile properties
                Console.WriteLine($"Display Name: {userProfileProperties.DisplayName}");
                Console.WriteLine($"Email: {userProfileProperties.Email}");
                Console.WriteLine($"Job Title: {userProfileProperties.Title}");
                // Access other properties as needed

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}
