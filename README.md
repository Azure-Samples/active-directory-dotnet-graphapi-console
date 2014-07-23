ConsoleApp-GraphAPI-DotNet
==========================

Console App using Graph Client Library Version 1.0

This console application is a .Net sample, using the Graph API Client library (Version 1.0) - it demonstrates common Read calls to the Graph API including Getting Users, Groups, Group Membership, Roles, Tenant information, Service Principals, Applications. The second part of the sample app demonstrates common Write/Update/Delete options on Users, Groups, and shows how to execute User License Assignment, updating a User's thumbnailPhoto and links, etc.  It also can read the contents from the signed-on user's mailbox.

The sample incorporates using the Active Directory Authentication Library (ADAL) for authentication. The first part of the console app is Read-only, and uses OAuth Client Credentials to authenticate against the Demo Company. The second part of the app has update operations, and requires User credentials (using the OAuth Authorization Code flow). Update operations are not permitted using this shared Demo company - to try these, you will need to update the application configuration to be used with your own Azure AD tenant, and will need to configure applications accordingly. When configuring this application to be used with your own tenant, you will need to configure two applications in the Azure Management Portal: one using OAuth Client Credentials, and a second one using OAuth Authorization Code flow (each with separate ClientIds (AppIds)) - to execute update operations, you will need to logon with an account that has Administrative permissions.  This app also demonstrates how to read the mailbox of the signed on user account using the Microsoft common consent framework  - applications can be granted access to Azure Active directory, as well as Microsoft Office365 applications including Exchange and SharePoint Online.  To configure the app:


Step 1: Clone or download this repository
From your shell or command line:

 git clone git@github.com:AzureADSamples/ConsoleApp-GraphAPI-DotNet.git 


Step 2: Run the sample in Visual Studio 2013
The sample app is preconfigured to read data from a Demonstration company (GraphDir1.onMicrosoft.com) in Azure AD. 
Run the sample application by selecting F5.  The second part of the app will require Admin credentials, you can simulate 
authentication using this demo user account: userName =  demoUser@graphDir1.onMicrosoft.com, password = graphDem0 
 However, this is only a user account and does not have administrative permissions to execute updates - therefore, you
will see "..unauthorized.." response errors when attempting any requests requiring admin permissions.  To see how updates
work, you will need to configure and use this sample with your own tenant - see the next step.


Step 3: Running this application with your own Azure Active Directory tenant

Register the Sample app for your own tenant

1.Sign in to the Azure management portal.

2.Click on Active Directory in the left hand nav.

3.Click the directory tenant where you wish to register the sample application.

4.Click the Applications tab.

5.In the drawer, click Add.

6.Click "Add an application my organization is developing".

7.Enter a friendly name for the application, for example "Console App for Azure AD", select "Web Application and/or Web API", and click next. 

8.For the Sign-on URL, enter a value (this is not used for the console app, so is only needed for this initial configuration):  "http://localhost"

9.For the App ID URI, enter "http://localhost".  Click the checkmark to complete the initial configuration.

10.While still in the Azure portal, click the Configure tab of your application.

11.Find the Client ID value and copy it aside, you will need this later when configuring your application.

12.Under the Keys section, select either a 1year or 2year key - the keyValue will be displayed after you save the configuration at the end - it will be displayed, and you should save this to a secure location. Note, that the key value is only displayed once, and you will not be able to retrieve it later.

13.Configure Permissions - under the "Permissions to other applications" section, you will configure permissions to access the Graph (Windows Azure Active Directory).
For "Windows Azure Active Directory" under the first permission column (Application Permission:1"), select "Read directory data". Notes: this configures the App to use OAuth Client Credentials, and have Read access permissions for the application. 

14.Select the Save button at the bottom of the screen - upon successful configuration, your Key value should now be displayed - please copy and store this value in a secure location.

15.You will need to update the program.cs of this Application project with the updated values. From Visual Studio, open the project and program.cs file, find and update the string values of "clientId" and "clientSecret" with the Client ID and key values from Azure management portal. Update your tenant name for the authString value (e.g. contoso.onMicrosoft.com).  Update the tenantId value for the string tenantId, with your tenantId.  Note: your tenantId can be discovered by opening the following metadata.xml document: https://login.windows.net/GraphDir1.onmicrosoft.com/FederationMetadata/2007-06/FederationMetadata.xml  - replace "graphDir1.onMicrosoft.com", with your tenant's domain value (any domain that is owned by the tenant will work).  The tenantId is a guid, that is part of the sts URL, returned in the first xml node's sts url ("EntityDescriptor"): e.g. "https://sts.windows.net/<tenantIdvalue>"

16.Now Configure a 2nd application object to run the update portion of this app: return to the Azure Management Portal's Application Page, select "Add" from the bottom, seelect "Add an Application my Organization is Developing", Supply an Application name, and make sure to select "Native Client Application", supply a redirect Uri (e.g. "https://localhost").  Select "configure" from the top tab - under "permissions to other applications" select the DelegatedPermissions:1 drop down menu for the Graph (Windows Azure Active Directory), and select "Access Your organization's directory".  This application will also attempt to read the signed-on user's Mailbox contents from Exchange Online - to enable this, add an additional permission: select "Office365 Exchange Online" and from the DeletagePermissions:1 drop down, select "Read users mail (preview)". Copy the Client ID value - this will be used to configure program.cs next - save the Application configuration.   
Select SAVE on the bottom of the screen.

17.Open the program.cs file, and find the "redirectUri" string value, and replace it with "https://localhost" (or the value your configured for the ReplyURL). Also replace the "clientIdForUserAuthn" with the client ID value from the previous step.

18.Build and run your application - you will need to authenticate with valid tenant administrator credentials for your company when you run the application (required for the Create/Update/delete operations).
