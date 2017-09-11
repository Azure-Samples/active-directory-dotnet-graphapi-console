using System;
using System.Threading.Tasks;
using Microsoft.Azure.ActiveDirectory.GraphClient;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace GraphConsoleAppV3
{
    internal class AuthenticationHelper
    {
        public static string TokenForUser;
        public static string TokenForApplication;

        /// <summary>
        /// Get Active Directory Client for Application.
        /// </summary>
        /// <returns>ActiveDirectoryClient for Application.</returns>
        public static ActiveDirectoryClient GetActiveDirectoryClientAsApplication()
        {
            Uri servicePointUri = new Uri(GlobalConstants.ResourceUrl);
            Uri serviceRoot = new Uri(servicePointUri, GlobalConstants.TenantId);
            ActiveDirectoryClient activeDirectoryClient = new ActiveDirectoryClient(serviceRoot,
                async () => await AcquireTokenAsyncForApplication());
            return activeDirectoryClient;
        }

        /// <summary>
        /// Async task to acquire token for Application.
        /// </summary>
        /// <returns>Async Token for application.</returns>
        public static async Task<string> AcquireTokenAsyncForApplication()
        {
            return await GetTokenForApplication();
        }

        /// <summary>
        /// Get Token for Application.
        /// </summary>
        /// <returns>Token for application.</returns>
        public static async Task<string> GetTokenForApplication()
        {
            if (TokenForApplication == null)
            {
                AuthenticationContext authenticationContext = new AuthenticationContext(
                    AppModeConstants.AuthString,
                    false);

                // Configuration for OAuth client credentials 
                if (string.IsNullOrEmpty(AppModeConstants.ClientSecret))
                {
                    Program.WriteError(
                        "Client secret not set. Please follow the steps in the README to generate a client secret.");
                }
                else
                {
                    ClientCredential clientCred = new ClientCredential(
                        GlobalConstants.ClientId,
                        AppModeConstants.ClientSecret);
                    AuthenticationResult authenticationResult =
                        await authenticationContext.AcquireTokenAsync(GlobalConstants.ResourceUrl, clientCred);
                    TokenForApplication = authenticationResult.AccessToken;
                }
            }
            return TokenForApplication;
        }

        /// <summary>
        /// Get Active Directory Client for User.
        /// </summary>
        /// <returns>ActiveDirectoryClient for User.</returns>
        public static ActiveDirectoryClient GetActiveDirectoryClientAsUser()
        {
            Uri servicePointUri = new Uri(GlobalConstants.ResourceUrl);
            Uri serviceRoot = new Uri(servicePointUri, GlobalConstants.TenantId);
            ActiveDirectoryClient activeDirectoryClient = new ActiveDirectoryClient(serviceRoot,
                async () => await AcquireTokenAsyncForUser());
            return activeDirectoryClient;
        }

        /// <summary>
        /// Async task to acquire token for User.
        /// </summary>
        /// <returns>Token for user.</returns>
        public static async Task<string> AcquireTokenAsyncForUser()
        {
            return await GetTokenForUser();
        }

        /// <summary>
        /// Get Token for User.
        /// </summary>
        /// <returns>Token for user.</returns>
        public static async Task<string> GetTokenForUser()
        {
            if (TokenForUser == null)
            {
                var redirectUri = new Uri("https://localhost");
                AuthenticationContext authenticationContext = new AuthenticationContext(UserModeConstants.AuthString, false);
                AuthenticationResult userAuthnResult = await authenticationContext.AcquireTokenAsync(GlobalConstants.ResourceUrl,
                    GlobalConstants.ClientId, redirectUri, new PlatformParameters(PromptBehavior.RefreshSession));
                TokenForUser = userAuthnResult.AccessToken;
                Console.WriteLine("\n Welcome " + userAuthnResult.UserInfo.GivenName + " " +
                                  userAuthnResult.UserInfo.FamilyName);
            }
            return TokenForUser;
        }

    }
}
