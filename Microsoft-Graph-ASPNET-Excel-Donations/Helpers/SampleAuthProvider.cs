/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

using Microsoft.Identity.Client;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.OpenIdConnect;
using Microsoft_Graph_ASPNET_Excel_Donations.TokenStorage;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web;
using System;

namespace Microsoft_Graph_ASPNET_Excel_Donations.Helpers
{
    public sealed class SampleAuthProvider : IAuthProvider
    {

        // Properties used to get and manage an access token.
        private string redirectUri = ConfigurationManager.AppSettings["ida:RedirectUri"];
        private string appId = ConfigurationManager.AppSettings["ida:AppId"];
        private string appSecret = ConfigurationManager.AppSettings["ida:AppSecret"];
        private string scopes = ConfigurationManager.AppSettings["ida:GraphScopes"];
        private SessionTokenCache tokenCache { get; set; }

        private static readonly SampleAuthProvider instance = new SampleAuthProvider();
        private SampleAuthProvider() { }

        public static SampleAuthProvider Instance
        {
            get
            {
                return instance;
            }
        }

        // Gets an access token. First tries to get the token from the token cache.
        public async Task<string> GetUserAccessTokenAsync()
        {
            string signedInUserID = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
            HttpContextWrapper httpContext = new HttpContextWrapper(HttpContext.Current);
            TokenCache userTokenCache = new SessionTokenCache(signedInUserID, httpContext).GetMsalCacheInstance();
            //var cachedItems = tokenCache.ReadItems(appId); // see what's in the cache

            ConfidentialClientApplication cca = new ConfidentialClientApplication(
                appId,
                redirectUri,
                new ClientCredential(appSecret),
                userTokenCache,
                null);

            try
            {
                var accounts = await cca.GetAccountsAsync();
                AuthenticationResult result = await cca.AcquireTokenSilentAsync(scopes.Split(new char[] { ' ' }), accounts.First());
                return result.AccessToken;
            }

            // Unable to retrieve the access token silently.
            catch (Exception)
            {
                HttpContext.Current.Request.GetOwinContext().Authentication.Challenge(
                    new AuthenticationProperties() { RedirectUri = "/" },
                    OpenIdConnectAuthenticationDefaults.AuthenticationType);

                throw new Exception("Caller needs to authenticate.");
            }
        }
    }
}