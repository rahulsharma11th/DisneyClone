using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using System.Net.Http;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace SharePointFunctionApp
{
    public static class GetSharePointSiteDetails
    {
        [FunctionName("GetSharePointSiteDetails")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            string siteUrl = Environment.GetEnvironmentVariable("SharePointSiteUrl");
            string tenantId = Environment.GetEnvironmentVariable("TenantId");
            string clientId = Environment.GetEnvironmentVariable("SharePointClientId");
            string clientSecret = Environment.GetEnvironmentVariable("SharePointClientSecret");

            string accessToken = await GetAppOnlyAccessToken(tenantId, clientId, clientSecret, $"https://{tenantId}.sharepoint.com");
            if (string.IsNullOrEmpty(accessToken))
            {
                return new BadRequestObjectResult("Unable to obtain access token.");
            }

            try
            {
                using (ClientContext context = new ClientContext(siteUrl))
                {
                    context.ExecutingWebRequest += (sender, e) =>
                    {
                        e.WebRequestExecutor.WebRequest.Headers["Authorization"] = "Bearer " + accessToken;
                    };

                    // Perform operations with SharePoint Online
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

        private static async Task<string> GetAppOnlyAccessToken(string tenantId, string clientId, string clientSecret, string resource)
        {
            using (var client = new HttpClient())
            {
                var values = new Dictionary<string, string>
                {
                    {"grant_type", "client_credentials"},
                    {"client_id", clientId},
                    {"client_secret", clientSecret},
                    {"resource", resource}
                };

                var content = new FormUrlEncodedContent(values);
                var response = await client.PostAsync($"https://accounts.accesscontrol.windows.net/{tenantId}/tokens/OAuth/2", content);
                var responseString = await response.Content.ReadAsStringAsync();
                var tokenResponse = System.Text.Json.JsonSerializer.Deserialize<TokenResponse>(responseString);
                return tokenResponse?.AccessToken;
            }
        }

        private class TokenResponse
        {
            [JsonPropertyName("access_token")]
            public string AccessToken { get; set; }
        }
    }
}
