# What to do with EWS Managed API PowerShell scripts that use Basic Authentication #

The EWS Managed API has been a good client-side library that has been around for a very long time and consumed in a lot of different PowerShell scripts over the years (in a number of differing ways). For the most part most of these implementations will be using Basic authentication (or NTLM/Kerberos if it's being used against onPrem). With the big switch off for basic Authentication coming in October there is a good chance that these old (or even something newly written) that has been running in the background somewhere and has become either forgotten or just neglected will stop working.


## How do you know if its using the EWS Managed API   ##

If a script is using the EWS Managed API it will have a line in the code that loads the Microsoft.Exchange.WebServices.dll. Depending on the version your using it may look like

    $dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
    [void][Reflection.Assembly]::LoadFile($dllpath)

or

    Import-Module -Name "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"

or the method I've used for a few years

    function Invoke-LoadEWSManagedAPI {
    param( 
    )  
    Begin {
        if (Test-Path ($script:ModuleRoot + "/Microsoft.Exchange.WebServices.dll")) {
            Import-Module ($script:ModuleRoot + "/Microsoft.Exchange.WebServices.dll")
            $Script:EWSDLL = $script:ModuleRoot + "/Microsoft.Exchange.WebServices.dll"
            write-verbose ("Using EWS dll from Local Directory")
        }
        else {

        
            ## Load Managed API dll  
            ###CHECK FOR EWS MANAGED API, IF PRESENT IMPORT THE HIGHEST VERSION EWS DLL, ELSE EXIT
            $EWSDLL = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")
            if (Test-Path $EWSDLL) {
                Import-Module $EWSDLL
                $Script:EWSDLL = $EWSDLL 
            }
            else {
                "$(get-date -format yyyyMMddHHmmss):"
                "This script requires the EWS Managed API 1.2 or later."
                "Please download and install the current version of the EWS Managed API from"
                "http://go.microsoft.com/fwlink/?LinkId=255472"
                ""
                "Exiting Script."
                exit
            } 
        }
    }
    }


or the fanciest method I've seen used in [https://github.com/OfficeDev/O365-InvestigationTooling/blob/master/Get-AllTenantRulesAndForms.ps1](https://github.com/OfficeDev/O365-InvestigationTooling/blob/master/Get-AllTenantRulesAndForms.ps1) where they store a compressed version of the dll in the script file.

## How do you know if its using Basic Authentication   ##

If it has a line like

    $Service.Credentials = New-Object System.Net.NetworkCredential($PSCredential.UserName.ToString(),$PSCredential.GetNetworkCredential().password.ToString())

or

    $service.Credentials = New-Object System.Net.NetworkCredential($username,$password)

Then it is using Basic authentication

## How do you know if it using OAuth ##

if its using OAuth you should see something like


    $OAuthCredentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($AccessToken)
    $service.Credentials = $OAuthCredentials

if that the case your all good and don't need this article.

## I'm using an App Pasword isn't that Modern Authentcation

No this is still Basic authentication, and this feature will disappear along with Basic authentication

## What to do next ##

The first thing to do is ask do you really need this script anymore? eg is there a better more reliable way of achieving what it’s doing. For most of these scripts the answer is generally yes and the reasons they are still being done this way are a lot more complicated and usually nothing to do with the underlying technology. 

## Unattended Scripts ##

If you have a script that runs unattended you may want to consider moving to the Client Credentials OAuth flow [https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-client-creds-grant-flow](https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-client-creds-grant-flow) and using Application polices to lock down Mailbox access to only the resources that are needed [https://docs.microsoft.com/en-us/graph/auth-limit-mailbox-access](https://docs.microsoft.com/en-us/graph/auth-limit-mailbox-access). This removes the need to have a userName and Password to run the script which is something you really want to do from a security point of view.


## I don't want to change much I just need to upgrade it or getting it working   ##

What you shouldn't do in this instance is try to re-enable basic authentication, while it may seem like a good idea, it's a terrible idea an also it will soon be impossible once Microsoft decommissioned it. 

If you want to upgrade your script to use Modern Auth here's a bunch of different posts from different people on how to do it.(these are all good methods)

[https://morgantechspace.com/2022/03/connect-ews-api-with-modern-authentication-using-powershell.html](https://morgantechspace.com/2022/03/connect-ews-api-with-modern-authentication-using-powershell.html)

[https://ingogegenwarth.wordpress.com/2018/08/02/ews-and-oauth/](https://ingogegenwarth.wordpress.com/2018/08/02/ews-and-oauth/)

[https://gsexdev.blogspot.com/2019/10/using-msal-microsoft-authentication.html](https://gsexdev.blogspot.com/2019/10/using-msal-microsoft-authentication.html)

Some example code in PowerShell if you were using an Endpoint in Azure (Based upon the three earlier blog posts) would look like this

     $TLS12Protocol = [System.Net.SecurityProtocolType] 'Ssl3 , Tls12'
     [System.Net.ServicePointManager]::SecurityProtocol = $TLS12Protocol

     #Provide your Office 365 Tenant Domain Name or Tenant Id
     $TenantId = "somedomain.onmicrosoft.com"

     #Provide Application (client) Id of your app
     $AppClientId="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

     #Provide Application client secret key
     $ClientSecret ="MySecretFromBackEndWebApp"

     $ID=[Microsoft.Exchange.WebServices.Data.ImpersonatedUserId]::new()
     $id.Id='user@contoso.com'
     $service.ImpersonatedUserId=$ID

     $RequestBody = @{client_id=$AppClientId;client_secret=$ClientSecret;grant_type="client_credentials";scope=”https://outlook.office365.com/.default”;}
     $OAuthResponse = Invoke-RestMethod -Method Post -Uri “https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token” -Body $RequestBody
     $AccessToken = $OAuthResponse.access_token
     
This would replace your PowerShell code for Basic Authentication to Modern Authentication

## Scripts that run past 1 hour ##

If you have a script that is running for a long time eg it might be processing every mailbox in a very large tenant you do need to be careful that AccessToken's do expire. So if you modify your script using any of the above links it may appear to work fine in testing but fail at a random point after 1 hour. So consider modifying your script to track the token expiration and refresh. 

## Dll Substitution to convert Basic to OAuth ##

One possible method instead of modifying your script to do the oAuth heavy lifting is because the EWS Managed API is open source you can create your own port of this code and implement the necessary Authentication code in your own custom version of the dll. I've created a branch of the EWS Managed API to demonstrate one possible example of this [https://github.com/gscales/ews-managed-api/tree/Force-Basic-ToOAuth](https://github.com/gscales/ews-managed-api/tree/Force-Basic-ToOAuth).With this version I've updated the .net framework the library uses to 4.72 and then added MSAL as a dependency so the MSAL code can be used directly in the EWS Managed API method. I've then modified some of the core EWS Managed API methods so it will force basic credentials to oAuth. For oAuth to work you still need an Application registration and TenantId information (and possibly a redirectURL) these are handled by the DLL reading them from a configuration file called Microsoft.Exchange.WebServices.OauthMod.dll.config eg

    <?xml version="1.0" encoding="utf-8" ?>
    <configuration>
      <appSettings>
    <add key="ClientId" value="xxxx-52b3-4102-aeff-aad2292ab01c" />
    <add key="oAuthMode" value="ROPC" />
    <add key="TenantId" value="xx8db77e-65e0-4fc3-b967-b14d5057375b" />
      </appSettings>
    </configuration>

Then when the code in the DLL runs to connect to EWS it forces the basic Authentication credentials to OAuth credentials using either the ROPC oauth flow (which means the script can continue to run unattended) or it can also do an Interactive authentication. One of the benefits of this method is that because it uses MSAL it supports token refresh and the only line of code that you need to change in the DLL is the line that imports the Managed API dll.

