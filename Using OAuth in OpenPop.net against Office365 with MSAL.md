## Using OAuth in OpenPop.net against Office365 with MSAL ##

If you are still using POP3 in any application to access email on Office365 then you want to pay imediate attention to the fact that Basic Authentication is going to be switched off for good (regardless of whether your using it or not) in October see [https://techcommunity.microsoft.com/t5/exchange-team-blog/basic-authentication-and-exchange-online-september-2021-update/ba-p/2772210](https://techcommunity.microsoft.com/t5/exchange-team-blog/basic-authentication-and-exchange-online-september-2021-update/ba-p/2772210) (Any they really really mean it this time)

So what should you do ? If you are using POP3 you should probably stop (it was great in 1988 but the world and security has moved on since then) and look at using the Microsoft Graph instead. However if thats not possible then you need to look at the libraries you are using in your code to access Pop3 and update the code to support modern authentication. One of the most popular way (based on nuget downloads) of implmenting POP3 in a C# apps is to use OpenPop.net [https://www.nuget.org/packages/OpenPop.NET/ ](https://www.nuget.org/packages/OpenPop.NET/). This libary doesn't support OAuth offically yet but the underlying POP3 protocol like SMTP,IMAP is pretty simple (which is another reason you shouldn't be using it from a security standpoint) so modifying this (or any simular libary) isn't a very hard thing to do. (A quick note if you have used MimeKit [https://github.com/jstedfast/MailKit](https://github.com/jstedfast/MailKit) which is my prefered library for this and many other things, this already support Modern Auth so you tasks is much easier see [http://www.mimekit.net/docs/html/T_MailKit_Security_SaslMechanismOAuth2.htm](http://www.mimekit.net/docs/html/T_MailKit_Security_SaslMechanismOAuth2.htm))

# Authentication Provider  #

If you are going to update the authentication in your application what you use to do the OAuth Authentication/Token acquisition part is probably the hardest decision/change you have to make. Eg your authentication code can be as simple as a few HttpClient requests up to a fully robust authentication libary like [MSAL](https://docs.microsoft.com/en-us/azure/active-directory/develop/msal-overview) which is the one I would recommend as its going to work in most identity topologies (which can get complicated around different federation providers etc). MSAL is what I'll be using in the examples in this doc.

# What you need to change in OpenPop.net   #

Because OpenPop.net doesn't currently support oAuth but is open source you can easly modify the source code to suit you needs. Generally doing this isn't a good idea as your can loose track of upstream updates but given the age of POP3 and the number of updates to this library it shouldn't really be that much of an issue. As an example I created a fork of OpenPop.net [https://github.com/gscales/hpop](https://github.com/gscales/hpop) and added oAuth support to this.

the two changes I made to this library where to add a new Authetication Method for XOAUTH2 in the Pop3client authenticaiton method

    	public void Authenticate(string username, string password, AuthenticationMethod authenticationMethod)
		......
		case AuthenticationMethod.XOAUTH2:
			AuthenticateUsingXOAUTH2(password);
			break;
And a method that sends the Authentication

    	private void AuthenticateUsingXOAUTH2(string saslXoAuthToken)
		{
			SendCommand("AUTH XOAUTH2");
			SendCommand(saslXoAuthToken);
			// Authentication was successful if no exceptions thrown before getting here
		}


# Putting in all together #

An example of using MSAL to get an access token for POP3 Authentication and then format that as a sasl token and use that in my OpenPop fork looks like



        NetworkCredential networkCredential = new NetworkCredential("blah@blah.onmicrosoft.com", "blah");
        PublicClientApplicationBuilder pcaConfig = PublicClientApplicationBuilder
           .Create("d64799fe-dfb2-480b-a3be-a7a5a0bdaf32").WithTenantId("eb8db77e-65e0-4fc3-b967-b14d5057375b");
        var app = pcaConfig.Build();
        var tokenResult = app.AcquireTokenByUsernamePassword(new string[] { "https://outlook.office.com/POP.AccessAsUser.All" }, networkCredential.UserName, networkCredential.SecurePassword).ExecuteAsync().GetAwaiter().GetResult();
        var saslformatedToken = Convert.ToBase64String(ASCIIEncoding.ASCII.GetBytes("user=" + networkCredential.UserName + (char)1 + "auth=Bearer " + tokenResult.AccessToken + (char)1 + (char)1));
        var client = new Pop3Client();
        client.Connect("outlook.office365.com", 995, true);
        client.Authenticate(networkCredential.UserName, saslformatedToken, AuthenticationMethod.XOAUTH2);
        int messageCount = client.GetMessageCount();
        var lastMessage = client.GetMessage(messageCount);            
      

