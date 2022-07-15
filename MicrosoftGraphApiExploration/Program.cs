using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Identity.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace MicrosoftGraphApiExploration
{
    class Program
    {
        private static GraphServiceClient graphServiceClient;

        public Program()
        {
            AuthenticationConfig config = AuthenticationConfig.ReadFromJsonFile("appsettings.json");

            // Even if this is a console application here, a daemon application is a confidential client application
            IConfidentialClientApplication app = ConfidentialClientApplicationBuilder.Create(config.ClientId)
                .WithClientSecret(config.ClientSecret)
                .WithAuthority(new Uri(config.Authority))
                .Build();

            app.AddInMemoryTokenCache();

            // With client credentials flows the scopes is ALWAYS of the shape "resource/.default", as the 
            // application permissions need to be set statically (in the portal or by PowerShell), and then granted by
            // a tenant administrator. 
            string[] scopes = new string[] { $"{config.ApiUrl}.default" }; // Generates a scope -> "https://graph.microsoft.com/.default"

            // Prepare an authenticated MS Graph SDK client
            graphServiceClient = GetAuthenticatedGraphClient(app, scopes);
        }
        static void Main(string[] args)
        {
            try
            {
                var program = new Program();
                program.UpdateUserMobileNumber().GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
                Console.ResetColor();
            }

            Console.WriteLine("Press any key to exit");
            Console.ReadKey();
        }


        /// <summary>
        /// The following example shows how to initialize the MS Graph SDK
        /// </summary>
        /// <param name="app"></param>
        /// <param name="scopes"></param>
        /// <returns></returns>
        private static async Task CallMSGraphUsingGraphSDK(IConfidentialClientApplication app, string[] scopes)
        {
            // Prepare an authenticated MS Graph SDK client
            GraphServiceClient graphServiceClient = GetAuthenticatedGraphClient(app, scopes);


            List<User> allUsers = new List<User>();

            try
            {

                IGraphServiceUsersCollectionPage users = await graphServiceClient.Users.Request().GetAsync();
                Console.WriteLine($"Found {users.Count()} users in the tenant");
            }
            catch (ServiceException e)
            {
                Console.WriteLine("We could not retrieve the user's list: " + $"{e}");
            }

        }

        private async Task UpdateUserMobileNumber()
        {
            var user = new User
            {
                BusinessPhones = new List<String>()
               {
                 "BusinessPhone"
               },
                OfficeLocation = "18/2111",
                MobilePhone = "MobilePhone"
            };

            var userUpdate = await graphServiceClient.Users["userid"]
                 .Request()
                 .UpdateAsync(user);


        }

        /// <summary>
        /// An example of how to authenticate the Microsoft Graph SDK using the MSAL library
        /// </summary>
        /// <returns></returns>
        private static GraphServiceClient GetAuthenticatedGraphClient(IConfidentialClientApplication app, string[] scopes)
        {

            GraphServiceClient graphServiceClient =
                    new GraphServiceClient("https://graph.microsoft.com/V1.0/", new DelegateAuthenticationProvider(async (requestMessage) =>
                    {
                        // Retrieve an access token for Microsoft Graph (gets a fresh token if needed).
                        AuthenticationResult result = await app.AcquireTokenForClient(scopes)
                            .ExecuteAsync();

                        // Add the access token in the Authorization header of the API request.
                        requestMessage.Headers.Authorization =
                            new AuthenticationHeaderValue("Bearer", result.AccessToken);
                    }));

            return graphServiceClient;
        }
    }
}
