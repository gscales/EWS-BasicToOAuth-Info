# Phython Office365 / Exchange Online IMAP sample using MSAL Imaplib and client credentials flow #

While its not recommneded to use IMAP4 anymore for connecting to Exchange i was struggling to find a complete end-end sample for connecting using the client credentials flow, msal and the Imaplib so I put together the following example and doco from a number of different sources. 

# Pre-Requisites #

This section is important and is different from EWS and Graph and where a lot of people get stuck. 

**Step 1** is create your applicaiton registration and consent to that registration as per [https://learn.microsoft.com/en-us/exchange/client-developer/legacy-protocols/how-to-authenticate-an-imap-pop-smtp-application-by-using-oauth#use-client-credentials-grant-flow-to-authenticate-imap-and-pop-connections](https://learn.microsoft.com/en-us/exchange/client-developer/legacy-protocols/how-to-authenticate-an-imap-pop-smtp-application-by-using-oauth#use-client-credentials-grant-flow-to-authenticate-imap-and-pop-connections). If you have problems finding the correct permission you can just modify the manifest directly and then consent to the permission in the Portal. The manifest permission for IMAP for the client credentials flow should be

	"requiredResourceAccess": [
		{
			"resourceAppId": "00000002-0000-0ff1-ce00-000000000000",
			"resourceAccess": [
				{
					"id": "5e5addcd-3e8d-4e90-baf5-964efab2b20a",
					"type": "Role"
				}
			]
		}
	],

**Step 2** This is important and is missed by the a lot TLDR with IMAP unlike EWS and Graph you must 

1. Register service principals in Exchange [https://learn.microsoft.com/en-us/exchange/client-developer/legacy-protocols/how-to-authenticate-an-imap-pop-smtp-application-by-using-oauth#register-service-principals-in-exchange](https://learn.microsoft.com/en-us/exchange/client-developer/legacy-protocols/how-to-authenticate-an-imap-pop-smtp-application-by-using-oauth#register-service-principals-in-exchange)
2. Explictly Grant Access to the Mailboxes your applicaiton will use eg 

`Add-MailboxPermission -Identity "john.smith@contoso.com" -User <SERVICE_PRINCIPAL_ID> -AccessRights FullAccess`

This last step is important and is different from EWS and Graph where the client credentials flow gives you access to all mailboxes in an organization on Exchange Online. This shouldn't be a problem for existing IMAP applications where some form of delegate access would have been nessasary. If your reading this and building a new script DONT use IMAP use the Graph eg [https://learn.microsoft.com/en-us/graph/tutorials/python?tabs=aad&tutorial-step=](https://learn.microsoft.com/en-us/graph/tutorials/python?tabs=aad&tutorial-step=5)5.

#Libaries 

To use this script in phython you need to install the MSAL library [https://github.com/AzureAD/microsoft-authentication-library-for-python](https://github.com/AzureAD/microsoft-authentication-library-for-python) eg

    pip install msal


#Script Login and Read First Email using Client Credentials
    import sys  
    import base64
    import json
    import logging
    import imaplib
    import msal
    import email
    
    config = {
    "authority": "https://login.microsoftonline.com/eb8db77e-65e0-4fc3-b967-....",
    "client_id": "18bb3888-dad0-4997-96b1-......",
    "scope": ["https://outlook.office.com/.default"],
    "secret": ".....",
    "tenant-id": "eb8db77e-65e0-4fc3-b967-...."
    }
    
    mailboxToAccess = 'user@domain.onmicrosoft.com'
    
    app = msal.ConfidentialClientApplication(config['client_id'], authority=config['authority'],
     client_credential=config['secret'])
    result = app.acquire_token_silent(config["scope"], account=None)
    
    def GenerateOAuth2String(username, access_token):
      auth_string = 'user=%s\1auth=Bearer %s\1\1' % (username, access_token)
      return auth_string
    
    if not result:
    logging.info("No suitable token exists in cache. Let's get a new one from AAD.")
    result = app.acquire_token_for_client(scopes=config["scope"])
    
    if "access_token" in result:
    user = mailboxToAccess
    server = 'outlook.office365.com'
    conn = imaplib.IMAP4_SSL(server)  
    conn.debug = 4
    conn.authenticate('XOAUTH2', lambda x: GenerateOAuth2String(user, result['access_token']))
    messages = conn.select("INBOX")
    print(messages[0])
    res, firstMessage = conn.fetch('1', '(RFC822)')
    msg = email.message_from_bytes(firstMessage[0][1])
    print(res)
    print(msg['Date'])
    print(msg['subject'])
    
    else:
    print(result.get("error"))
    print(result.get("error_description"))
    print(result.get("correlation_id"))  # You may need this when reporting a bug



   
  
