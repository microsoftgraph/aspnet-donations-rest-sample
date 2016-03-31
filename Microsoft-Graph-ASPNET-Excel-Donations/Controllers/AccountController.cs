using System.Web;
using System.Web.Mvc;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.Cookies;
using Microsoft.Owin.Security.OpenIdConnect;
using Microsoft_Graph_ASPNET_Excel_Donations.TokenStorage;

namespace Microsoft_Graph_ASPNET_Excel_Donations.Controllers
{
    public class AccountController : Controller
    {
        public void SignIn()
        {

                // Signal OWIN to send an authorization request to Azure
                HttpContext.GetOwinContext().Authentication.Challenge(
                  new AuthenticationProperties { RedirectUri = "/Donation" },
                  OpenIdConnectAuthenticationDefaults.AuthenticationType);
            

        }

        public void SignOut()
        {
            if (Request.IsAuthenticated)
            {
                // Get the user's token cache and clear it
                string userObjId = System.Security.Claims.ClaimsPrincipal.Current
                  .FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;

                SessionTokenCache tokenCache = new SessionTokenCache(userObjId, HttpContext);
                tokenCache.Clear();
            }
            // Send an OpenID Connect sign-out request. 
            HttpContext.GetOwinContext().Authentication.SignOut(
              CookieAuthenticationDefaults.AuthenticationType);
            Response.Redirect("/");
        }
    }
}