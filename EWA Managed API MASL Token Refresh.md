# Using MSAL in the EWS Managed API and doing auto token expiration and renewal in Delegate and Client Credential Azure oAuth Flows

## Synopsis 

A lot of applications use the EWS Managed API to make interfacing with EWS easier and with the depreciation of Basic Authentication are looking for ways to migrate their code to use OAuth. EWS provides an OAuthCredentials class but this class does not provide any call-back mechanisms to allow for token caching, expiration and renewal as provided by the MSAL (or ADAL) authentication libraries.

# Solution
The EWS Managed API is an open-source library https://github.com/OfficeDev/ews-managed-api so it can be forked and modified as needed. This post provides one possible C# implementation of how to solve the above problem, but other methods are possible.


# Implmentation

By default, the EWS Managed API supports oAuth authentication but doesn’t provide a method to manage token expiration and refresh. Eg the default OAuthCredentials class has one constructor that allows you to pass in the AccessToken eg

`exchangeService.Credentials = new OAuthCredentials(TokenResult.AccessToken);`

That token will then be used each time when the PrepareWebRequest override method is executed in the OAuthCredentials class. Eg

        internal override void PrepareWebRequest(IEwsHttpWebRequest request)
        {
            base.PrepareWebRequest(request);

            if (this.token != null)
            {
                request.Headers.Remove(HttpRequestHeader.Authorization);
                request.Headers.Add(HttpRequestHeader.Authorization, this.token);
            }
            else
            {
                request.Credentials = this.credentials;
            }
        }

Because the OAuthCredentials class is a sealed class in the EWS Managed API you can’t create your own implementation or override in a consuming project.  

If you have an application where a Token expiration of an hour will cause the application to fail, a method to check Token expiration and refresh will be required in the code. The Microsoft MSAL library supports token expiration, caching and refresh, so provides an optimal solution but it requires the modification of the EWS Managed API library code which is available on GitHub https://github.com/OfficeDev/ews-managed-api . There are multiple ways you could modify the library to support the refresh and the method you use should reflect that which fits the coding style and best practices in your Organization. The following example creates a new abstract credentials class within the EWS Managed API that you can then create your own implementation of in a consuming project. This mean that both the MSAL,ADAL or any type of authentication library can be supported as you will have access to the PrepareWebRequest flow. A modified fork of the EWS Managed API that demonstrates this is available https://github.com/gscales/ews-managed-api 

A very simple implementation of this that supports both Delegate and Application permissions and shows a simple console app example of using both Delegate and Application permissions and also allows for the support of Hybrid Modern Authentication.



	namespace BasicToOAuthImplmentation
	{
    public class MSALDelegateTokenClass : Microsoft.Exchange.WebServices.Data.Credentials.CustomTokenCredentials
    {
        private const string GraphScope = "https://graph.microsoft.com/User.Read";
        private string ClientId { get; set; }
        private string TenantId { get; set; }
        private string RedirectUri { get; set; }

        private string EWSAuthScope { get; set; } = "https://outlook.office.com/EWS.AccessAsUser.All";
        public MSALDelegateTokenClass(string clientId, string tenantid, string redirectUri)
        {
            ClientId = clientId;
            TenantId = tenantid;
            RedirectUri = redirectUri;
        }
        public IPublicClientApplication app { get; set; }
        public override string GetCustomToken()
        {
       
            if (app == null)
            {
                PublicClientApplicationBuilder pcaConfig = PublicClientApplicationBuilder.Create(ClientId).WithTenantId(TenantId);
                app = pcaConfig.WithRedirectUri(RedirectUri).Build();
            }
            var accounts = app.GetAccountsAsync().GetAwaiter().GetResult();
            AuthenticationResult result = null;
            try
            {
                result = app.AcquireTokenSilent(new[] { EWSAuthScope }, accounts.FirstOrDefault())
                                  .ExecuteAsync().GetAwaiter().GetResult();
            }
            catch (MsalUiRequiredException ex)
            {
                // A MsalUiRequiredException happened on AcquireTokenSilent.
                // This indicates you need to call AcquireTokenInteractive to acquire a token
                try
                {
                    result = app.AcquireTokenInteractive(new[] { EWSAuthScope })
                                      .ExecuteAsync().GetAwaiter().GetResult();
                }
                catch (MsalException msalex)
                {
                    Console.WriteLine(msalex.Message);
                    //
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            return "Bearer " + result.AccessToken;
        }

        public string GetUserName()
        {
            if (app == null)
            {
                var Token = GetCustomToken();
            }
            var accounts = app.GetAccountsAsync().GetAwaiter().GetResult();
            return accounts.FirstOrDefault().Username;
        }

        public string GetEmailAddress()
        {

            if (app == null)
            {
                var token = GetCustomToken();
            }
            var accounts = app.GetAccountsAsync().GetAwaiter().GetResult();      
            var graphToken = app.AcquireTokenSilent(new[] { GraphScope }, accounts.FirstOrDefault())
                                  .ExecuteAsync().GetAwaiter().GetResult();
            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", graphToken.AccessToken);
                var profileResult =  client.GetAsync("https://graph.microsoft.com/v1.0/me?$select=mail").GetAwaiter().GetResult();
                dynamic profile = JObject.Parse(profileResult.Content.ReadAsStringAsync().GetAwaiter().GetResult());
                return profile.mail;
            }
        }

        public string SetAuthenticationScope(string emailAddress)
        {
           
            string autodiscoverv2Endpoint = $"https://outlook.office365.com/autodiscover/autodiscover.json/v1.0/{emailAddress}?Protocol=EWS";
            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (compatible; AcmeInc/1.0)");
                dynamic adResponse = JObject.Parse(client.GetAsync(autodiscoverv2Endpoint).GetAwaiter().GetResult().Content.ReadAsStringAsync().GetAwaiter().GetResult());
                if(adResponse.Url != null)
                {
                    EWSAuthScope = "https://" + new Uri(adResponse.Url.ToString()).Host + "/EWS.AccessAsUser.All";
                }
                return adResponse.Url;
            }
            
        }

    }

    public class MSALAppTokenClass : Microsoft.Exchange.WebServices.Data.Credentials.CustomTokenCredentials
    {
        private string ClientId { get; set; }
        private string TenantId { get; set; }
        private SecureString ClientSecret { get; set; }

        private string EWSAuthScope { get; set; } = "https://outlook.office.com/.default";

        public MSALAppTokenClass(string clientId, string tenantid, SecureString clientSecret)
        {
            ClientId = clientId;
            TenantId = tenantid;
            ClientSecret = clientSecret;
        }
        public IConfidentialClientApplication app { get; set; }
        public override string GetCustomToken()
        {
            

            if (app == null)
            {
                app = ConfidentialClientApplicationBuilder.Create(ClientId)
                 .WithClientSecret(new System.Net.NetworkCredential(string.Empty, ClientSecret).Password)
                 .WithTenantId(TenantId)                 
                 .Build();              

            }           
            AuthenticationResult result = null;
            try
            {
                result = app.AcquireTokenForClient(new[] { EWSAuthScope }).ExecuteAsync().Result;
            }
            catch (Exception ex)
            {
                throw;
            }
            return "Bearer " + result.AccessToken;
        }

        public string SetAuthenticationScope(string emailAddress)
        {

            string autodiscoverv2Endpoint = $"https://outlook.office365.com/autodiscover/autodiscover.json/v1.0/{emailAddress}?Protocol=EWS";
            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (compatible; AcmeInc/1.0)");
                dynamic adResponse = JObject.Parse(client.GetAsync(autodiscoverv2Endpoint).GetAwaiter().GetResult().Content.ReadAsStringAsync().GetAwaiter().GetResult());
                if (adResponse.Url != null)
                {
                    EWSAuthScope = "https://" + new Uri(adResponse.Url.ToString()).Host + "/.default";
                }
                return adResponse.Url;
            }

        }


    }
    internal class Program
    {
        
        static void Main(string[] args)
        {
            TestDelegatePermissions();
            List<string> mailboxesToAccess = new List<string>{
                "user1@domain.com",
                "user2@domain.com"
            };
            TestApplicationPermissions(mailboxesToAccess,new NetworkCredential("", "...").SecurePassword);
        }

        static void TestDelegatePermissions()
        {
            MSALDelegateTokenClass delegateCreds = new MSALDelegateTokenClass("9d5d77a6-fe09-473e-8931-958f15f1a96b", "1c3a18bf-da31-4f6c-a404-2c06c9cf5ae4", "msal9d5d77a6-fe09-473e-8931-958f15f1a96b://auth");
            var emailAddress = delegateCreds.GetEmailAddress();
            var ewsURL = delegateCreds.SetAuthenticationScope(emailAddress);
            ExchangeService exchangeService = new ExchangeService();
            exchangeService.Credentials = delegateCreds;
            // For Autodiscover V1 use
            // exchangeService.AutodiscoverUrl(emailAddress, RedirectionUrlValidationCallback);
            exchangeService.Url = new Uri(ewsURL);
            exchangeService.HttpHeaders.Add("X-AnchorMailbox", emailAddress);
            var InboxFolder = Folder.Bind(exchangeService, WellKnownFolderName.Inbox);
            Console.WriteLine("Done");
        }

        static void TestApplicationPermissions(List<String> mailboxesToAccess,SecureString clientSecret)
        {
            ExchangeService exchangeService = new ExchangeService();
            MSALAppTokenClass mSALAppToken = new MSALAppTokenClass("c957131d-b228-494e-b1b8-32b6605fadb9", "1c3a18bf-da31-4f6c-a404-2c06c9cf5ae4", clientSecret);
            exchangeService.Credentials = mSALAppToken;
            foreach(String mailboxToAccess in mailboxesToAccess)
            {
                exchangeService.Url = new Uri(mSALAppToken.SetAuthenticationScope(mailboxToAccess));
                exchangeService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, mailboxToAccess);
                exchangeService.HttpHeaders.Remove("X-AnchorMailbox");
                exchangeService.HttpHeaders.Add("X-AnchorMailbox", mailboxToAccess);
                var InboxFolder = Folder.Bind(exchangeService, WellKnownFolderName.Inbox);
                Console.WriteLine("Mailbox " + mailboxToAccess);
            }
            Console.WriteLine("Done");
        }
        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;
            Uri redirectionUri = new Uri(redirectionUrl);
            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }
    }
}`
