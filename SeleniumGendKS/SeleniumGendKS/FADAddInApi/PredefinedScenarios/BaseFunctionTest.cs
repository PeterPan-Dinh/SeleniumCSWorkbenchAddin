using Microsoft.Identity.Client;
using NUnit.Framework;
using SeleniumGendKS.Pages;
using System.Configuration;
using System.Xml.Linq;
using System.Xml.XPath;

namespace SeleniumGendKS.FADAddInApi.PredefinedScenarios
{
    [TestFixture]
    internal class BaseFunctionTest
    {
        // Initiate variables (to get msal.idtoken)
        internal static readonly string? email = LoginPage.email; 
        internal static readonly string? password = LoginPage.password;
        internal static readonly string? clientId = LoginPage.clientId; 
        internal static readonly string? tenantId = LoginPage.tenantId;
        internal static string authority = $"https://login.microsoftonline.com/{tenantId}";
        internal static string[] scopes = new[] { "Files.Read.All", "openid", "profile", "User.Read", "email" };
        internal static readonly string? redirectUri = LoginPage.redirectUri;
        internal static IPublicClientApplication? app;

        // Login to get msal.idtoken - MSAL (Microsoft Authentication Library)
        internal static async Task<AuthenticationResult> LoginAsync()
        {
            app = PublicClientApplicationBuilder.Create(clientId)
                 .WithAuthority(authority)
                 .WithRedirectUri(redirectUri) // Redirect URI for desktop/mobile apps
                 .Build();
            AuthenticationResult? result = await app.AcquireTokenByUsernamePassword(scopes, email, password).ExecuteAsync();
            return result;
        }

        // get value from .config file
        internal static readonly XDocument xdoc = XDocument.Load(@"Config\Config.xml");
        internal static readonly string? bearerToken = xdoc.XPathSelectElement("config/tokens").Attribute("bearerToken").Value;
        internal static readonly string? msalIdtoken = LoginAsync().Result.IdToken;

        [SetUp]
        public void Setup()
        {
            string? getMSALIdtoken = msalIdtoken;
            if (getMSALIdtoken == null || getMSALIdtoken == "")
                getMSALIdtoken = xdoc.XPathSelectElement("config/tokens").Attribute("msalIdtoken").Value;
        }
    }
}
