using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using System.Security;

namespace SharePointFunctionApp
{
    public static class GetUserProfile
    {
        [FunctionName("GetUserProfile")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            string siteUrl = Environment.GetEnvironmentVariable("SharePointSiteUrl");
            string clientId = Environment.GetEnvironmentVariable("SharePointClientId");
            string clientSecret = Environment.GetEnvironmentVariable("SharePointClientSecret");

            string accountName = req.Query["accountName"];
            if (string.IsNullOrEmpty(accountName))
            {
                return new BadRequestObjectResult("Please pass an accountName on the query string or in the request body");
            }

            try
            {
                using (ClientContext context = new ClientContext(siteUrl))
                {
                    SecureString secureString = new SecureString();
                    foreach (char c in clientSecret)
                    {
                        secureString.AppendChar(c);
                    }

                    context.AuthenticationMode = ClientAuthenticationMode.AppOnly;
                    context.Credentials = new SharePointOnlineCredentials(clientId, secureString);

                    // Get the user profile
                    PeopleManager peopleManager = new PeopleManager(context);
                    PersonProperties personProperties = peopleManager.GetPropertiesFor(accountName);
                    context.Load(personProperties);
                    await context.ExecuteQueryAsync();

                    // Extracting and returning user profile properties
                    return new OkObjectResult($"User profile for {accountName}: {personProperties.DisplayName}, {personProperties.Email}");
                }
            }
            catch (Exception ex)
            {
                log.LogError($"Exception occurred: {ex.Message}");
                return new BadRequestObjectResult($"Error occurred: {ex.Message}");
            }
        }
    }
}
