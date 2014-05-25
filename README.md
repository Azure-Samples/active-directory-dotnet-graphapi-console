ConsoleApp-GraphAPI-DotNet
==========================

Console App using Graph Client Library Version 1.0

This console application is a .Net sample, using the Graph API Client library (Version 1.0) - it demonstrates common Read calls 
to the Graph API including Getting Users, Groups, Group Membership, Roles, Tenant information, Service Principals, Applications.
The second part of the sample app demonstrates common Write/Update/Delete options on Users, Groups, and shows how to execute 
User License Assignement, updating a User's thumbnailPhoto and links.

The sample incorporates using the Active Directory Authentication Library (ADAL) for authentication.  The first part of the console app is 
Read-only, and uses OAuth Client Credentials to authenticate against the Demo Company.  The second part has Update operations, and 
requires User credentials (using the OAuth Authorization Code flow).  Update operations are not permitted using this Demo company -
to try these, you will need to update the application configuration to be used with your own Azure AD tenant, and will need
to configure applications accordingly.  To do this:



Step 1: Clone or download this repository
From your shell or command line:

 git clone git@github.com:AzureADSamples/ConsoleApp-GraphAPI-DotNet.git 


Step 2: Run the sample in Visual Studio 2013
The sample app is preconfigured to read data from a Demonstration company (GraphDir1.onMicrosoft.com) in Azure AD. 
Run the sample application by selecting F5.  The second part of the app will require Admin credentials, you can simulate 
authentication using thid demo user account: demoUser@graphDir1.onMicrosoft.com graphDem0 
However, this is only a user account and does not have administrative permissions to execute updates - therefore, you
will see "..unauthorized.." response errors when attemping any requests requiring admin permissions.  To see how updates
work, you will need to configure and use this sample with your own tenant - see the next step.


Step 3: Running this application with your Azure Active Directory tenant

Register the Sample app for your own tenant
1.Sign in to the Azure management portal.
2.Click on Active Directory in the left hand nav.
3.Click the directory tenant where you wish to register the sample application.
4.Click the Applications tab.
5.In the drawer, click Add.
6.Click "Add an application my organization is developing".
7.Enter a friendly name for the application, for example "Console App for Azure AD", select "Web Application and/or Web API", and click next.
8.For the sign-on URL, enter the base URL for the sample, which is by default  https://localhost:44322 .
9.For the App ID URI, enter  https://<your_tenant_name>/WebAppGraphAPI , replacing  <your_tenant_name>  with the domain name of your Azure AD tenant. For Example "https://contoso.com/WebAppGraphAPI". Click OK to complete the registration.
10.While still in the Azure portal, click the Configure tab of your application.
11.Find the Client ID value and copy it aside, you will need this later when configuring your application.
12.In the Reply URL, add the reply URL address used to return the authorization code returned during Authorization code flow. For example: "https://localhost:44322/"
13.Under the Keys section, select either a 1year or 2year key - the keyValue will be displayed after you save the configuration at the end - it will be displayed, and you should save this to a secure location. Note, that the key value is only displayed once, and you will not be able to retrieve it later.


14.Configure Permissions - under the "Permissions to other applications" section, select application "Windows Azure Active Directory" (this is the Graph API), and under the second permission (Delegated permissions), select "Access your organization's directory" and "Enable sign-on and read users' profiles". The 2nd column (Application permission) is not needed for this demo app. Notes: the permission "Access your organization's directory" allows the application to access your organization's directory on behalf of the signed-in user - this is a delegation permission and must be consented by the Administrator for webApps (such as this demo app). The permission "Enable sign-on and read users' profiles" allows users to sign in to the application with their organizational accounts and lets the application read the profiles of signed-in users, such as their email address and contact information - this is a delegation permission, and can be consented to by the user. The other permissions, "Read Directory data" and "Read and write Directory data", are Delegation and Application Permissions, which only the Administrator can grant consent to.


15.Selct the Save button at the bottom of the screen - upon sucessful configuration, your Key value should now be displayed - please copy and store this value in a secure location.


16.You will need to update the webconfig file of this Application project. From Visual Studio, open the web.config file, and under the section, modify "ida:ClientId" and "ida:AppKey" and " with the values from the previous steps. Also update the "ida:Tenant" with your Azure AD Tenant's domain name e.g. Contoso.onMicrosoft.com, (or Contoso.com if that domain is owned by your tenant).

17.In  web.config  add this line in the  <system.web>  section:  <sessionState timeout="525600" /> . This increases the ASP.Net session state timeout to it's maximum value so that access tokens and refresh tokens cache in session state aren't cleared after the default timeout of 20 minutes.
18.Build and run your application - you will need to authenticate with valid user credentials for your company when you run the application.
