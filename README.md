---
services: active-directory
platforms: dotnet
author: dstrockis
---

# Call the Azure AD Graph API from a native client

Console App using Graph Client Library Version 2.0 (updated to use version 2.1.1)

> **NOTE**: Azure AD Graph API functionality is also available through [**Microsoft Graph**](https://graph.microsft.com), a unified API that also includes APIs from other Microsoft services like Outlook, OneDrive, SharePoint, and many more, all accessed through a single endpoint with a single access token. It is strongly recommended that developers use [**Microsoft Graph**](https://graph.microsft.com) in preference to Azure AD Graph API and client library. Microsoft Graph has a [.Net Client Library](https://www.nuget.org/packages/Microsoft.Graph), is also available as [open source SDK code](https://github.com/microsoftgraph/msgraph-sdk-dotnet), with UWP and ASP.Net samples. The most similar sample app to this console app sample is the [**UWP snippets sample**](https://github.com/microsoftgraph/uwp-csharp-snippets-sample). Unlike the Azure AD Graph client library (only available for .Net), Microsoft Graph client libraries are available on many platforms and languages and are all open source.  Full details are available [here](https://graph.microsoft.io/en-us/code-samples-and-sdks).


This console application is a .Net sample, using the Graph API Client library (Version 2.0). It demonstrates common calls to the Graph API including Getting Users, Groups, Group Membership, Roles, Tenant information, Service Principals, Applications and Domains. The sample app demonstrates both delegated authentication (user mode, requiring a user to sign in to the app and access graph API) and application authentication (app mode, where the app access is made using the application's identity without a user needing to be present).  When running in user mode, some operations that the sample performs will **only** be possible if the signed-in user is a company administrator.

The sample uses the Active Directory Authentication Library (ADAL) for authentication. Out of the box, the console sample is configured to run against a demo tenant, with sample user credentials provided (as described in Step 2 below).  This user *only* has read-only access to the demo tenant. To see all operations working, you'll need to this application to be used with your own tenant. This requires configuring two applications in the Azure Portal: one using *application permissions* or OAuth Client Credentials, and a second one using user *delegated permissions* - to execute update operations, you will need to sign-in with an account that has Administrative permissions. Configuring the console sample against your own tenant is described in Step 3 below).


## Step 1: Clone or download this repository
From your shell or command line:

`git clone https://github.com/Azure-Samples/active-directory-dotnet-graphapi-console.git`


## Step 2: Run the sample in Visual Studio 2013 or 2015
The sample app is preconfigured to read data from a Demonstration company (GraphDir1.onMicrosoft.com) in Azure AD. 
Run the sample application by selecting F5, selecting to run in both modes (user and app mode).  The app will require Admin credentials to perform *all* operations, and you can simulate authentication using a demo user account: userName =  demoUser@graphDir1.onMicrosoft.com, password = graphDem0 

However, this is only a user account and does not have administrative permissions to execute updates - therefore, you
will see "..unauthorized.." response errors when attempting any requests requiring admin permissions.  To see how updates
work, you will need to configure and use this sample with your own tenant - see the next step.

## Step 3: Running this application with your own Azure Active Directory tenant
Register the Sample app for your own tenant 

1. Sign in to the [Azure Portal](https://portal.azure.com) using an account from your own Azure Active Directory tenant.

2. Type in **App registrations** in the search bar.  You'll now register 2 applications for this sample.

3. In the **App registrations blade**, click **Add**, and enter a friendly name for the application, for example **"Console App for Azure AD"**, select **Native**, add use **https://localhost/** as the Redirect URI. Click **Create**.

4. You should now see your newly created app in the app list.  Click on the app to see further settings.

5. Click **Required permissions** (in the **Settings** blade).  You'll now see *Windows Azure Active Directory* listed. Select this to configure additional permissions. Under **Enable permissions**, select **Access the directory as the signed-in user**, and then click **Save**.

6. Finally, under **Settings**, **Properties** find and copy the Application ID value for later *user mode* use.

7. Now configure the console sample code to use the user mode app you just registered above. From Visual Studio, open the project and Constants.cs file, and replace the **ClientId** in **UserModeConstants** with the Application ID value from the previous step.

8. Now we'll need to configure a second application for the application mode portion of the console app. Follow steps 1-3, however instead, at step 3, select **Web app / API**, and enter "http://localhost/" for the **Sign-on URL**. (NOTE: this is not used for the console app, so is only needed for this initial configuration.)

9. You should now see your newly created app in the app list.  Click on the app to see further settings.

10. Click **Keys** (in the **Settings** blade), and enter a key description and select a key duration from one of the 3 options.  The **key value** will be displayed after you click **Save**. Save this to a secure location - you'll need it later for the sample app code. NOTE: The key value is only displayed once, and you will not be able to retrieve it later.

11. Click **Required permissions** (in the **Settings** blade).  You'll now see *Windows Azure Active Directory* listed. Select this to configure additional permissions. Under **Enable permissions**, select *Read Directory Data* from the list of **Application permissions**.  Selecting an application permission will mean that your application can use the OAuth client credential flow to call the Graph API (without requiring a user). Click **Save**.

12. Finally, under **Settings**, **Properties** find and copy the Application ID value for later *app mode use*.

13. Now to configure the console sample code to use the app mode app you just registered above. From Visual Studio, open the project and Constants.cs file. In the **AppModeConstants** class: 
+ Update the string values of **ClientId** and **ClientSecret** with the Application ID and key value (steps 12 and 10 respectively).
+ Update your **TenantName** to represent your tenant name (e.g. contoso.onMicrosoft.com). 
+ Update the **TenantId** value. Your tenant ID can be discovered by opening the following metadata.xml document: https://login.microsoftonline.com/GraphDir1.onmicrosoft.com/FederationMetadata/2007-06/FederationMetadata.xml  - replace "graphDir1.onMicrosoft.com", with your tenant's domain value (any domain that is owned by your tenant will work).  The tenantId is a guid, that is part of the sts URL, returned in the first xml node's sts url ("EntityDescriptor"): e.g. "https://sts.windows.net/<tenantIdvalue>"

14. Build and run your application.  
+ You might run into some "missing assembly reference?" errors when you build. Make sure NuGet package restore is enabled, and that the packages in packages.config are installed. Sometimes, Visual Studio doesn't immediately find the package, so try building again. If all else fails, you can add the references manually.
+ When you run it, select user mode the first time. You will need to authenticate with valid tenant administrator credentials for your company when you run the application (required for the Create/Update/Delete operations), and consent the first time you use the sample app.
+ If you want to run the console app in app mode (which is somewhat contrived, since this really should be run as a native client), you'll need to force consent manually, beforehand. Here, an admin user will need to consent.  You can force consent by opening a browser, and going to the following URL, replacing **\<tenantId\>** with your tenantId, and **\<app-mode-application-id\>** with your Application ID for your app mode application:
   ```http
   https://login.microsoftonline.com/<tenantId>/oauth2/authorize?
   client_id=<app-mode-application-id>
   &response_type=code
   &redirect_uri=http%3A%2F%2Flocalhost%2F&response_mode=query
   &resource=https%3A%2F%2Fgraph.windows.net%2F&state=12345
   ```
After signing in (if not already signed in), click the **Accept** in the consent page.  You can then close the browser.  Now that you've pre-consented, you can try running the console sample in app mode.
