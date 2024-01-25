using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using System.Net.Http;
using System.Collections.Generic;
using System;
using System.Text.Json.Serialization;

namespace SharePointProfileFunction
{
    public static class GetUserProfile
    {
        [FunctionName("GetUserProfile")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            var sharePointAuth = new SharePointAuth(
                Environment.GetEnvironmentVariable("TenantId"),
                Environment.GetEnvironmentVariable("AdminUrl"),
                Environment.GetEnvironmentVariable("UserProfileClientId"),
                Environment.GetEnvironmentVariable("UserProfileClientSecret")
            );
            var accessToken = await sharePointAuth.GetAccessToken();

            if (!string.IsNullOrEmpty(accessToken))
            {
                string accountName = req.Query["accountName"];
                var userProfile = await sharePointAuth.GetUserProfileForUser(accessToken, accountName);
                return new OkObjectResult(userProfile);
            }
            else
            {
                return new BadRequestObjectResult("Failed to get access token.");
            }
        }

        private class SharePointAuth
        {
            private readonly string TenantId;
            private readonly string AdminUrl;
            private readonly string UserProfileClientId;
            private readonly string UserProfileClientSecret;
            private readonly string TokenEndpoint;

            public SharePointAuth(string tenantId, string adminUrl, string userProfileClientId, string userProfileClientSecret)
            {
                TenantId = tenantId;
                AdminUrl = adminUrl;
                UserProfileClientId = userProfileClientId;
                UserProfileClientSecret = userProfileClientSecret;
                TokenEndpoint = $"https://accounts.accesscontrol.windows.net/{TenantId}/tokens/OAuth/2";
            }

            public async Task<string> GetAccessToken()
            {
                using (var client = new HttpClient())
                {
                    var values = new Dictionary<string, string>
                    {
                        { "grant_type", "client_credentials" },
                        { "client_id", UserProfileClientId },
                        { "client_secret", UserProfileClientSecret },
                        { "resource", AdminUrl }
                    };

                    var response = await client.PostAsJsonAsync(TokenEndpoint, values);
                    if (!response.IsSuccessStatusCode)
                    {
                        Console.WriteLine("Error: " + response.StatusCode);
                        return null;
                    }

                    var tokenResponse = await response.Content.ReadFromJsonAsync<TokenResponse>();
                    return tokenResponse?.AccessToken;
                }
            }

            public async Task<string> GetUserProfileForUser(string accessToken, string accountName)
            {
                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);

                    string encodedAccountName = Uri.EscapeDataString(accountName);
                    string userProfileUrl = $"{AdminUrl}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='{encodedAccountName}'";
                    var response = await client.GetAsync(userProfileUrl);

                    if (!response.IsSuccessStatusCode)
                    {
                        Console.WriteLine("Error retrieving user profile: " + response.StatusCode);
                        return null;
                    }

                    var userProfile = await response.Content.ReadAsStringAsync();
                    return userProfile;
                }
            }

            private class TokenResponse
            {
                [JsonPropertyName("access_token")]
                public string AccessToken { get; set; }
            }
        }
    }
}
