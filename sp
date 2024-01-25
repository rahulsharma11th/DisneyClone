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

                    // Example SharePoint operation: Load web title
                    Web web = context.Web;
                    context.Load(web);
                    await context.ExecuteQueryAsync();

                    return new OkObjectResult($"SharePoint site title: {web.Title}");
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
