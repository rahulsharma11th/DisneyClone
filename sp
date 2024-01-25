using System;
using System.Net.Http;
using System.Net.Http.Json;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace SharePointUserProfile
{
    class Program
    {
        static async Task Main(string[] args)
        {
            var sharePointAuth = new SharePointAuth(
                "your-tenant-id",
                "https://your-admin-url",
                "user-profile-client-id",
                "user-profile-client-secret"
            );
            var accessToken = await sharePointAuth.GetAccessToken();

            if (!string.IsNullOrEmpty(accessToken))
            {
                Console.WriteLine("Access Token: " + accessToken);

                // Specify the user's account name here
                string accountName = "i:0#.f|membership|user_email@example.com";
                var userProfile = await sharePointAuth.GetUserProfileForUser(accessToken, accountName);
                Console.WriteLine("User Profile for " + accountName + ": " + userProfile);
            }
            else
            {
                Console.WriteLine("Failed to get access token.");
            }
        }
    }

    public class SharePointAuth
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


//////////////////////////////////////////////
{
  "SharePoint": {
    "TenantId": "your-tenant-id",
    "AdminUrl": "https://your-admin-url",
    "UserProfileClientId": "user-profile-client-id",
    "UserProfileClientSecret": "user-profile-client-secret"
  }
}
appsettings.json


using System;
using System.Net.Http;
using System.Net.Http.Json;
using System.Threading.Tasks;
using System.Collections.Generic;
using Microsoft.Extensions.Configuration;
using System.Text.Json.Serialization;
using System.IO;

namespace SharePointUserProfile
{
    class Program
    {
        static async Task Main(string[] args)
        {
            var configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();

            var sharePointAuth = new SharePointAuth(configuration);
            var accessToken = await sharePointAuth.GetAccessToken();

            if (!string.IsNullOrEmpty(accessToken))
            {
                Console.WriteLine("Access Token: " + accessToken);

                // Specify the user's account name here
                string accountName = "i:0#.f|membership|user_email@example.com";
                var userProfile = await sharePointAuth.GetUserProfileForUser(accessToken, accountName);
                Console.WriteLine("User Profile for " + accountName + ": " + userProfile);
            }
            else
            {
                Console.WriteLine("Failed to get access token.");
            }
        }
    }

    public class SharePointAuth
    {
        private readonly string TenantId;
        private readonly string AdminUrl;
        private readonly string UserProfileClientId;
        private readonly string UserProfileClientSecret;
        private readonly string TokenEndpoint;

        public SharePointAuth(IConfiguration configuration)
        {
            TenantId = configuration["SharePoint:TenantId"];
            AdminUrl = configuration["SharePoint:AdminUrl"];
            UserProfileClientId = configuration["SharePoint:UserProfileClientId"];
            UserProfileClientSecret = configuration["SharePoint:UserProfileClientSecret"];
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
