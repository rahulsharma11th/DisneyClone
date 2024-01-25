using Microsoft.Office.Server;
using Microsoft.Office.Server.UserProfiles;
using Microsoft.SharePoint;
using System;

namespace SharePointUserProfileRetrieval
{
    class Program
    {
        static void Main(string[] args)
        {
            using (SPSite site = new SPSite("http://YourSharePointSiteURL"))
            {
                SPServiceContext context = SPServiceContext.GetContext(site);
                UserProfileManager profileManager = new UserProfileManager(context);

                string accountName = "domain\\username"; // Specify the account name
                if (profileManager.UserExists(accountName))
                {
                    UserProfile userProfile = profileManager.GetUserProfile(accountName);
                    
                    // Example: Get the user's work phone and department
                    Console.WriteLine("Work Phone: " + userProfile["WorkPhone"].Value);
                    Console.WriteLine("Department: " + userProfile["Department"].Value);
                }
                else
                {
                    Console.WriteLine("User profile not found.");
                }
            }
        }
    }
}
