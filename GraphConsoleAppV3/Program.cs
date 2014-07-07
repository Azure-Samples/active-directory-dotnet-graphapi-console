using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Xml;
using System.Net;
using Newtonsoft.Json;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Azure.ActiveDirectory.GraphClient;
using Microsoft.Azure.ActiveDirectory.GraphClient.ErrorHandling;

namespace GraphConsoleAppV3
{

    class Program
    {
        // Single-Threaded Apartment required for OAuth2 Authz Code flow (User Authn) to execute for this demo app
        [STAThread]
        static void Main()
        {
            // get OAuth token using Client Credentials
            string tenantName = "GraphDir1.onMicrosoft.com";
            string authString = "https://login.windows.net/" + tenantName;
     
            AuthenticationContext authenticationContext = new AuthenticationContext(authString,false);

            // Config for OAuth client credentials 
            string clientId = "118473c2-7619-46e3-a8e4-6da8d5f56e12";
            string clientSecret = "hOrJ0r0TZ4GQ3obp+vk3FZ7JBVP+TX353kNo6QwNq7Q=";
            ClientCredential clientCred = new ClientCredential(clientId, clientSecret);
            string resource = "https://graph.windows.net";
            string token;
            try
            {
                AuthenticationResult authenticationResult = authenticationContext.AcquireToken(resource, clientCred);
                token = authenticationResult.AccessToken;
            }
                
            catch (AuthenticationException ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Acquiring a token failed with the following error: {0}", ex.Message);
                if (ex.InnerException != null)
                {
                    //You should implement retry and back-off logic per the guidance given here:http://msdn.microsoft.com/en-us/library/dn168916.aspx
                    //InnerException Message will contain the HTTP error status codes mentioned in the link above
                    Console.WriteLine("Error detail: {0}", ex.InnerException.Message);
                }
                Console.ResetColor();
                Console.ReadKey();
                return;
            }

            
            // record start DateTime of execution
            string CurrentDateTime = DateTime.Now.ToUniversalTime().ToString();

            //*********************************************************************
            // setup Graph connection
            //*********************************************************************
            Guid ClientRequestId = Guid.NewGuid();
            GraphSettings graphSettings = new GraphSettings();
            graphSettings.ApiVersion = "2013-11-08";
            graphSettings.GraphDomainName = "graph.windows.net";
            GraphConnection graphConnection = new GraphConnection(token, ClientRequestId,graphSettings);
            VerifiedDomain initialDomain = new VerifiedDomain();
            VerifiedDomain defaultDomain = new VerifiedDomain();

            //*********************************************************************
            // Get Tenant Details
            // Note: update the string tenantId with your TenantId.
            // This can be retrieved from the login Federation Metadata end point:             
            // https://login.windows.net/GraphDir1.onmicrosoft.com/FederationMetadata/2007-06/FederationMetadata.xml
            //  Replace "GraphDir1.onMicrosoft.com" with any domain owned by your organization
            // The returned value from the first xml node "EntityDescriptor", will have a STS URL
            // containing your tenantId e.g. "https://sts.windows.net/4fd2b2f2-ea27-4fe5-a8f3-7b1a7c975f34/" is returned for GraphDir1.onMicrosoft.com
            //*********************************************************************
            string tenantId = "4fd2b2f2-ea27-4fe5-a8f3-7b1a7c975f34";
            GraphObject tenant = graphConnection.Get(typeof(TenantDetail), tenantId);
 
            if (tenant == null)
            {
                Console.WriteLine("Tenant not found");
            }
            else
            { 
                TenantDetail tenantDetail = (TenantDetail)tenant;
                Console.WriteLine("Tenant Display Name: " + tenantDetail.DisplayName);
            
                // Get the Tenant's Verified Domains 
                initialDomain = tenantDetail.VerifiedDomains.First(x => x.Initial.HasValue && x.Initial.Value);
                Console.WriteLine("Initial Domain Name: " + initialDomain.Name);
                defaultDomain = tenantDetail.VerifiedDomains.First(x => x.Default.HasValue && x.Default.Value);
                Console.WriteLine("Default Domain Name: " + defaultDomain.Name);
                foreach (string techContact in tenantDetail.TechnicalNotificationMails)
                {
                    Console.WriteLine("Tenant Tech Contact: " + techContact);
                }
            }
            //*********************************************************************
            // Demonstrate Getting a list of Users with paging (get 4 users), sort by displayName
            //*********************************************************************
            Console.WriteLine("\n Retrieving Users");
            
            FilterGenerator userFilter = new FilterGenerator();
            userFilter.Top = 4;
            userFilter.OrderByProperty = GraphProperty.DisplayName;
            PagedResults<User> users = graphConnection.List<User>(null, userFilter);
            foreach (User user in users.Results)
            {
                Console.WriteLine("UserObjectId: {0}  UPN: {1}", user.ObjectId, user.UserPrincipalName);
            }

            // if there are more users to retrieve, get the rest of the users, and specify maximum page size 999       
            do
            {
                userFilter.Top = 999;
                if (users.PageToken != null)
                {
                    users = graphConnection.List<User>(users.PageToken, userFilter);
                    Console.WriteLine("\n Next page of results");
                }
                             
                foreach (User user in users.Results)
                {
                    Console.WriteLine("DisplayName: {0}  UPN: {1}", user.DisplayName, user.UserPrincipalName);
                }
            } while (users.PageToken != null);

            // search for a single user by UPN
            string searchString = "adam@" + initialDomain.Name;
            FilterGenerator filter = new FilterGenerator();
            Expression filterExpression = ExpressionHelper.CreateEqualsExpression(typeof(User), GraphProperty.UserPrincipalName, searchString );
            filter.QueryFilter = filterExpression;
            

            User retrievedUser = new User();
            PagedResults<User> pagedUserResults = graphConnection.List<User>(null, filter);

            // should only find one user with the specified UPN
            if (pagedUserResults.Results.Count == 1)
            {
                retrievedUser = pagedUserResults.Results[0] as User;
            }
            else
            {
                Console.WriteLine("User not found {0}", searchString);
            }
            
            if (retrievedUser.UserPrincipalName != null)
            {
                Console.WriteLine("\n Found User: " + retrievedUser.DisplayName + "  UPN: " + retrievedUser.UserPrincipalName);

                // get the user's Manager
                    int count = 0;
                    PagedResults<GraphObject> managers = graphConnection.GetLinkedObjects(retrievedUser, LinkProperty.Manager, null);
                    foreach (GraphObject managerObject in managers.Results)
                    {
                      if (managerObject.ODataTypeName.Contains("User"))
                        {
                          User manager = (User)managers.Results[count];
                          Console.WriteLine(" Manager: {0}  UPN: {1}", manager.DisplayName, manager.UserPrincipalName);
                        }
                      count++;
                    }
                //*********************************************************************
                // get the user's Direct Reports
                //*********************************************************************
                int top = 99;
                PagedResults<GraphObject> directReportObjects = graphConnection.GetLinkedObjects(retrievedUser, LinkProperty.DirectReports, null, top);
                foreach (GraphObject graphObject in directReportObjects.Results)
                {
                    if (graphObject.ODataTypeName.Contains("User"))
                    {
                        User User = (User)graphObject;
                        Console.WriteLine(" DirectReport {0}: {1}  UPN: {2}", User.ObjectType, User.DisplayName, User.UserPrincipalName);
                    }

                    if (graphObject.ODataTypeName.Contains("Contact")) 
                    {
                        Contact Contact = (Contact)graphObject;
                        Console.WriteLine(" DirectReport {0}: {1}  Mail: {2} ", Contact.ObjectType, Contact.DisplayName, Contact.Mail);
                    }
                }
                //*********************************************************************
                // get a list of Group IDs that the user is a member of
                //*********************************************************************
                Console.WriteLine("\n {0} is a member of the following Groups (IDs)", retrievedUser.DisplayName);
                bool securityGroupsOnly = false;
                IList<string> usersGroupMembership = graphConnection.GetMemberGroups(retrievedUser, securityGroupsOnly);
                foreach (String groupId in usersGroupMembership)
                {
                    Console.WriteLine("Member of Group ID: "+ groupId);
                }


                //*********************************************************************
                // get the User's Group and Role membership, getting the complete objects
                //*********************************************************************
                PagedResults<GraphObject> memberOfObjects = graphConnection.GetLinkedObjects(retrievedUser, LinkProperty.MemberOf, null, top);
                foreach (GraphObject graphObject in memberOfObjects.Results)
                {
                    if (graphObject.ODataTypeName.Contains("Group"))
                    {
                        Group Group = (Group)graphObject;
                        Console.WriteLine(" Group: {0}  Description: {1}", Group.DisplayName, Group.Description);
                    }

                    if (graphObject.ODataTypeName.Contains("Role")) 
                    {
                        Role Role = (Role)graphObject;
                        Console.WriteLine(" Role: {0}  Description: {1}", Role.DisplayName, Role.Description);
                    }
                }
            }
            //*********************************************************************
            // People picker
            // Search for a user using text string "ad" match against userPrincipalName, proxyAddresses, displayName, giveName, surname   
            //*********************************************************************
            searchString = "ad";
            Console.WriteLine("\nSearching for any user with string {0} in UPN,ProxyAddresses,DisplayName,First or Last Name", searchString);
            
            FilterGenerator userMatchFilter = new FilterGenerator();
            userMatchFilter.Top = 19;
            Expression firstExpression = ExpressionHelper.CreateStartsWithExpression(typeof(User), GraphProperty.UserPrincipalName, searchString);
            Expression secondExpression = ExpressionHelper.CreateAnyExpression(typeof(User), GraphProperty.ProxyAddresses, "smtp:" + searchString);
            userMatchFilter.QueryFilter = ExpressionHelper.JoinExpressions(firstExpression, secondExpression, ExpressionType.Or);
            
            Expression thirdExpression = ExpressionHelper.CreateStartsWithExpression(typeof(User), GraphProperty.DisplayName, searchString);
            userMatchFilter.QueryFilter = ExpressionHelper.JoinExpressions(userMatchFilter.QueryFilter, thirdExpression, ExpressionType.Or);
            
            Expression fourthExpression = ExpressionHelper.CreateStartsWithExpression(typeof(User), GraphProperty.GivenName, searchString);
            userMatchFilter.QueryFilter = ExpressionHelper.JoinExpressions(userMatchFilter.QueryFilter, fourthExpression, ExpressionType.Or);

            Expression fifthExpression = ExpressionHelper.CreateStartsWithExpression(typeof(User), GraphProperty.Surname, searchString);
            userMatchFilter.QueryFilter = ExpressionHelper.JoinExpressions(userMatchFilter.QueryFilter, fifthExpression, ExpressionType.Or);

            PagedResults<User> serachResults = graphConnection.List<User>(null, userMatchFilter);

            if (serachResults.Results.Count > 0)
            {
                foreach (User User in serachResults.Results)
                {
                    Console.WriteLine("User DisplayName: {0}  UPN: {1}", User.DisplayName, User.UserPrincipalName);
                }
            }
            else
            {
                Console.WriteLine("User not found");
            }

            //*********************************************************************
            // Search for a group using a startsWith filter (displayName property)
            //*********************************************************************
            Group retrievedGroup = new Group();
            searchString = "Wash";
            filter.QueryFilter = ExpressionHelper.CreateStartsWithExpression(typeof(Group), GraphProperty.DisplayName, searchString);
            filter.Top = 99;

            PagedResults<Group> pagedGroupResults = graphConnection.List<Group>(null, filter);

            if (pagedGroupResults.Results.Count > 0)
            {
                retrievedGroup = pagedGroupResults.Results[0] as Group;
            }
            else
            {
                Console.WriteLine("Group Not Found");
            }

            if (retrievedGroup.ObjectId != null)
            {
                Console.WriteLine("\n Found Group: " + retrievedGroup.DisplayName + "  " + retrievedGroup.Description);

                //*********************************************************************
                // get the groups' membership using GetAllDirectLinks - 
                // Note this method retrieves ALL links in one request - please use this method with care - this
                // may return a very large number of objects
                //*********************************************************************

                GraphObject graphObj = (GraphObject)retrievedGroup;
               
                IList<GraphObject> members = graphConnection.GetAllDirectLinks(graphObj, LinkProperty.Members);
                if (members.Count > 0)
                {
                    Console.WriteLine("Group Membership");
                    foreach (GraphObject graphObject in members)
                    {
                        if (graphObject.ODataTypeName.Contains("User"))
                        {
                            User User = (User)graphObject;
                            Console.WriteLine("User DisplayName: {0}  UPN: {1}", User.DisplayName, User.UserPrincipalName);
                        }

                        if (graphObject.ODataTypeName.Contains("Group"))
                        {
                            Group Group = (Group)graphObject;
                            Console.WriteLine("Group DisplayName: {0}", Group.DisplayName);
                        }

                        if (graphObject.ODataTypeName.Contains("Contact"))
                        {
                            Contact Contact = (Contact)graphObject;
                            Console.WriteLine("Contact DisplayName: {0}", Contact.DisplayName);
                        }
                    }
                }
            }
            //*********************************************************************
            // Search for a Role by displayName
            //*********************************************************************
            searchString = "Company Administrator";
            filter.QueryFilter = ExpressionHelper.CreateStartsWithExpression(typeof(Role), GraphProperty.DisplayName, searchString);
            PagedResults<Role> pagedRoleResults = graphConnection.List<Role>(null, null);

            if (pagedRoleResults.Results.Count > 0)
            {
                foreach (GraphObject graphObject in pagedRoleResults.Results)
                {
                    Role role = graphObject as Role;
                    if (role.DisplayName == searchString.Trim())
                    {
                        Console.WriteLine("\n Found Role: {0} {1} {2} ", role.DisplayName, role.Description, role.ObjectId);
                    }
                }
            }
            else
            {
                Console.WriteLine("Role Not Found {0}",searchString);
            }

            //*********************************************************************
            // get the Service Principals
            //*********************************************************************
            filter.Top = 999;
            filter.QueryFilter = null;
            PagedResults<ServicePrincipal> servicePrincipals = new PagedResults<ServicePrincipal>();
            do
            {
                servicePrincipals = graphConnection.List<ServicePrincipal>(servicePrincipals.PageToken, filter);
                if (servicePrincipals != null)
                {
                    foreach (ServicePrincipal servicePrincipal in servicePrincipals.Results)
                    {
                        Console.WriteLine("Service Principal AppId: {0}  Name: {1}", servicePrincipal.AppId, servicePrincipal.DisplayName);
                    }
                }
            } while (servicePrincipals.PageToken != null);

            //*********************************************************************
            // get the  Application objects
            //*********************************************************************
            filter.Top = 999;
            PagedResults<Application> applications = new PagedResults<Application>();
            do
            {
                applications = graphConnection.List<Application>(applications.PageToken, filter);
                if (applications != null)
                {
                    foreach (Application application in applications.Results)
                    {
                        Console.WriteLine("Application AppId: {0}  Name: {1}", application.AppId, application.DisplayName);
                    }
                }
             }while (applications.PageToken != null);

            string targetAppId = applications.Results[0].ObjectId;


            //********************************************************************************************
            //  We'll now switch to Authenticating using OAuth Authorization Code Grant
            //  which includes user Authentication/Delegation
            //*********************************************************************************************
            var redirectUri = new Uri("https://localhost");
            string clientIdForUserAuthn = "66133929-66a4-4edc-aaee-13b04b03207d";
            AuthenticationResult userAuthnResult = null;
            try
            {
                userAuthnResult = authenticationContext.AcquireToken(resource, clientIdForUserAuthn, redirectUri, PromptBehavior.Always);
                token = userAuthnResult.AccessToken;                
                Console.WriteLine("\n Welcome " + userAuthnResult.UserInfo.GivenName + " " + userAuthnResult.UserInfo.FamilyName);
            }
            catch (AuthenticationException ex)
            {
                string message = ex.Message;
                if (ex.InnerException != null)
                    message += "InnerException : " + ex.InnerException.Message;
                Console.WriteLine(message);
                Console.ReadKey();
                return;
            }

            // re-establish Graph connection using the new token
            graphConnection = new GraphConnection(token, ClientRequestId, graphSettings);

            //*********************************************************************************************
            // Create a new User with a temp password
            //*********************************************************************************************
            User userToBeAdded = new User();
            userToBeAdded.DisplayName = "Sample App Demo User";
            userToBeAdded.UserPrincipalName = "SampleAppDemoUser@" + defaultDomain.Name;
            userToBeAdded.AccountEnabled = true;
            userToBeAdded.MailNickname = "SampleAppDemoUser";
            userToBeAdded.PasswordProfile = new PasswordProfile();
            userToBeAdded.PasswordProfile.Password = "TempP@ssw0rd!";
            userToBeAdded.PasswordProfile.ForceChangePasswordNextLogin = true;
            userToBeAdded.UsageLocation = "US";
            User newlyCreatedUser = new User();
            try 
            { 
              newlyCreatedUser = graphConnection.Add<User>(userToBeAdded);
              Console.WriteLine("\nNew User {0} was created", userToBeAdded.DisplayName);
            }
            catch (GraphException graphException)
            {
                Console.WriteLine("\nError creating new user " + graphException.ToString());
            }

            //*********************************************************************************************
            // update the newly created user's Password, PasswordPolicies and City
            //*********************************************************************************************
            if (newlyCreatedUser.ObjectId != null)
            {
                string userObjectId = newlyCreatedUser.ObjectId;

                // update User's city and reset their User's password
                User updateUser = graphConnection.Get<User>(userObjectId);
                updateUser.City = "Seattle";
                PasswordProfile passwordProfile = new PasswordProfile();
                passwordProfile.Password = "newP@ssw0rd!";
                passwordProfile.ForceChangePasswordNextLogin = false;
                updateUser.PasswordProfile = passwordProfile;
                updateUser.PasswordPolicies = "DisablePasswordExpiration, DisableStrongPassword";
                try
                {
                    graphConnection.Update(updateUser);
                    Console.WriteLine("\nUser {0} was updated", updateUser.DisplayName);
                }
                catch (GraphException graphException)
                {
                    Console.WriteLine("\nError Updating the user " + graphException.ToString());
                }


                //*********************************************************************************************
                // Add, then retrieve a thumbnailPhoto for the newly created user
                //*********************************************************************************************
                Bitmap thumbnailPhoto = new Bitmap(20, 20);
                thumbnailPhoto.SetPixel(5, 5, Color.Beige);
                thumbnailPhoto.SetPixel(5, 6, Color.Beige);
                thumbnailPhoto.SetPixel(6, 5, Color.Beige);
                thumbnailPhoto.SetPixel(6, 6, Color.Beige);

                using (MemoryStream ms = new MemoryStream())
                {
                    thumbnailPhoto.Save(ms, ImageFormat.Jpeg);
                    graphConnection.SetStreamProperty(newlyCreatedUser, GraphProperty.ThumbnailPhoto, ms, "image/jpeg");
                    //  graphConnection.SetStreamProperty(newlyCreatedUser, "thumbnailPhoto", ms, "image/jpeg");
                }

                using (Stream ms = graphConnection.GetStreamProperty(newlyCreatedUser, GraphProperty.ThumbnailPhoto, "image/jpeg"))
                {
                    Image jpegImage = Image.FromStream(ms);
                }



                //*********************************************************************************************
                // User License Assignment - assign EnterprisePack license to new user, and disable SharePoint service
                //   first get a list of Tenant's subscriptions and find the "Enterprisepack" one
                //   Enterprise Pack includes service Plans for ExchangeOnline, SharePointOnline and LyncOnline
                //   validate that Subscription is Enabled and there are enough units left to assign to users
                //*********************************************************************************************
                PagedResults<SubscribedSku> skus = graphConnection.List<SubscribedSku>(null, null);
                foreach (SubscribedSku sku in skus.Results)
                {
                    if (sku.SkuPartNumber == "ENTERPRISEPACK")
                        if ((sku.PrepaidUnits.Enabled.Value > sku.ConsumedUnits) && (sku.CapabilityStatus == "Enabled"))
                        {
                            // create addLicense object and assign the Enterprise Sku GUID to the skuId
                            // 
                            AssignedLicense addLicense = new AssignedLicense();
                            addLicense.SkuId = sku.SkuId.Value;

                            // find plan id of SharePoint Service Plan
                            foreach (ServicePlanInfo servicePlan in sku.ServicePlans)
                            {
                                if (servicePlan.ServicePlanName.Contains("SHAREPOINT"))
                                {
                                    addLicense.DisabledPlans.Add(servicePlan.ServicePlanId.Value);
                                    break;
                                }
                            }

                            IList<AssignedLicense> licensesToAdd = new AssignedLicense[] { addLicense };
                            IList<Guid> licensesToRemove = new Guid[] { };

                            // attempt to assign the license object to the new user 
                            try
                            {
                                graphConnection.AssignLicense(newlyCreatedUser, licensesToAdd, licensesToRemove);
                                Console.WriteLine("\n User {0} was assigned license {1}", newlyCreatedUser.DisplayName, addLicense.SkuId);
                            }
                            catch (GraphException graphException)
                            {
                                Console.WriteLine("\nLicense assingment failed " + graphException.ToString());
                            }

                        }
                }

                //*********************************************************************************************
                // Add User to the "WA" Group 
                //*********************************************************************************************
                if (retrievedGroup.ObjectId != null)
                {
                    try
                    {
                        graphConnection.AddLink(retrievedGroup, newlyCreatedUser, LinkProperty.Members);
                        Console.WriteLine("\nUser {0} was added to Group {1}", newlyCreatedUser.DisplayName, retrievedGroup.DisplayName);
                    }
                    catch (GraphException graphException)
                    {
                        Console.WriteLine("\nAdding user to group failed " + graphException.ToString());
                    }
                }

                //*********************************************************************************************
                // Create a new Group
                //*********************************************************************************************
                Group CaliforniaEmployees = new Group();
                CaliforniaEmployees.DisplayName = "California Employees";
                CaliforniaEmployees.Description = "Employees in the state of California";
                CaliforniaEmployees.MailNickname = "CalEmployees";
                CaliforniaEmployees.MailEnabled = false;
                CaliforniaEmployees.SecurityEnabled = true;
                Group newGroup = null;
                try
                {
                    newGroup = graphConnection.Add<Group>(CaliforniaEmployees);
                    Console.WriteLine("\nNew Group {0} was created", newGroup.DisplayName);
                }
                catch (GraphException graphException)
                {
                    Console.WriteLine("\nError creating new Group " + graphException.ToString());
                }

                //*********************************************************************************************
                // Add the new User member to the new Group
                //*********************************************************************************************
                if (newGroup.ObjectId != null)
                {
                    try
                    {
                        graphConnection.AddLink(newGroup, newlyCreatedUser, LinkProperty.Members);
                        Console.WriteLine("\nUser {0} was added to Group {1}", newlyCreatedUser.DisplayName, newGroup.DisplayName);
                    }
                    catch (GraphException graphException)
                    {
                        Console.WriteLine("\nAdding user to group failed " + graphException.ToString());
                    }
                }


                //*********************************************************************************************
                // Delete the user that we just created
                //*********************************************************************************************
                if (newlyCreatedUser.ObjectId != null)
                {
                    try
                    {
                        graphConnection.Delete(newlyCreatedUser);
                        Console.WriteLine("\nUser {0} was deleted", newlyCreatedUser.DisplayName);
                    }
                    catch (GraphException graphException)
                    {
                        Console.WriteLine("Deleting User failed" + graphException.ToString());
                    }
                }

                //*********************************************************************************************
                // Delete the Group that we just created
                //*********************************************************************************************
                if (newGroup.ObjectId != null)
                {
                    try
                    {
                        graphConnection.Delete(newGroup);
                        Console.WriteLine("\nGroup {0} was deleted", newGroup.DisplayName);
                    }
                    catch (GraphException graphException)
                    {
                        Console.WriteLine("Deleting Group failed" + graphException.ToString());
                    }
                }

            }

            //*********************************************************************************************
            // Get a list of Mobile Devices from tenant
            //*********************************************************************************************
            Console.WriteLine("\nGetting Devices");
            FilterGenerator deviceFilter = new FilterGenerator();
            deviceFilter.Top = 999;
            PagedResults<Device> devices = graphConnection.List<Device>(null, deviceFilter);
            foreach(Device device in devices.Results)
            {
                if (device.ObjectId !=null)
                {
                    Console.WriteLine("Device ID: {0}, Type: {1}", device.DeviceId, device.DeviceOSType);
                    foreach (GraphObject owner in device.RegisteredOwners)
                    {
                        Console.WriteLine("Device Owner ID: " + owner.ObjectId);
                    }
                }
            }

            //*********************************************************************************************
            // Create a new Application object
            //*********************************************************************************************
            Application appObject = new Application();
            appObject.DisplayName = "Test-Demo App";
            appObject.IdentifierUris.Add("https://localhost/demo/" + Guid.NewGuid().ToString());
            appObject.ReplyUrls.Add("https://localhost/demo");
            
            // created Keycredential object for the new App object
            KeyCredential KeyCredential = new KeyCredential();
            KeyCredential.StartDate = DateTime.UtcNow;
            KeyCredential.EndDate = DateTime.UtcNow.AddYears(1);
            KeyCredential.Type = "Symmetric";
            KeyCredential.Value = Convert.FromBase64String("g/TMLuxgzurjQ0Sal9wFEzpaX/sI0vBP3IBUE/H/NS4=");
            KeyCredential.Usage = "Verify";
            appObject.KeyCredentials.Add(KeyCredential);

            GraphObject newApp = null;
            try
            {
                newApp = graphConnection.Add(appObject);
                Console.WriteLine("New Application created: " + newApp.ObjectId);
            }
            catch (GraphException graphException)
            {
                Console.WriteLine("Application Creation execption: " + graphException.Message);
            }
            
            // Get the application object that was just created

            GraphObject app = graphConnection.Get(typeof(Application), newApp.ObjectId);
            Application retrievedApp = (Application)app;

            //*********************************************************************************************
            // create a new Service principal
            //*********************************************************************************************
            ServicePrincipal newServicePrincpal = new ServicePrincipal();
            newServicePrincpal.DisplayName = "Test-Demo App";
            newServicePrincpal.AccountEnabled = true;
            newServicePrincpal.AppId = retrievedApp.AppId;

            GraphObject newSP = null;
            try 
            {
                newSP = graphConnection.Add<ServicePrincipal>(newServicePrincpal);
            //    newSP = graphConnection.Add(newServicePrincpal);
                Console.WriteLine("New Service Principal created: " + newSP.ObjectId);
            }
            catch (GraphException graphException)
            {
                Console.WriteLine("Service Principal Creation execption: " + graphException.ToString());
            }

            //*********************************************************************************************
            // get all Permission Objects
            //*********************************************************************************************
            Console.WriteLine("\n Getting Permissions");
            filter.Top = 999;
            PagedResults<Permission> permissions = new PagedResults<Permission>();
            do 
            {
                    try
                    {
                        permissions = graphConnection.List<Permission>(permissions.PageToken, filter);
                    }
                    catch (GraphException graphException)
                    {
                        Console.WriteLine("Error: " + graphException.ToString());
                        break;
                    }
                               
                    foreach (Permission permission in permissions.Results)
                    {
                        Console.WriteLine("Permission: {0}  Name: {1}", permission.ClientId, permission.Scope);
                    }
                
            } while(permissions.PageToken != null);

            //*********************************************************************************************
            // Create new permission object
            //*********************************************************************************************
            Permission permissionObject = new Permission();
            permissionObject.ConsentType = "AllPrincipals";
            permissionObject.Scope = "user_impersonation";
            permissionObject.StartTime = DateTime.Now;
            permissionObject.ExpiryTime = (DateTime.Now).AddMonths(12);

            // resourceId is objectId of the resource, in this case objectId of AzureAd (Graph API)
            permissionObject.ResourceId = "dbf73c3e-e80b-495b-a82f-2f772bb0a417";
            
            //ClientId = objectId of servicePrincipal
            permissionObject.ClientId = newSP.ObjectId;

            GraphObject newPermission = null;
              try
              {
                newPermission = graphConnection.Add(permissionObject);
                Console.WriteLine("New Permission object created: " + newPermission.ObjectId);
              }
              catch (GraphException graphException)
              {
                Console.WriteLine("Permission Creation exception: " + graphException.ToString());
              }

            //*********************************************************************************************
            // Delete Application Objects
            //*********************************************************************************************

            if (retrievedApp.ObjectId != null)
            {
                try
                {
                    graphConnection.Delete(retrievedApp);
                    Console.WriteLine("Deleting Application object: " + retrievedApp.ObjectId);
                }
                catch (GraphException graphException)
                {
                    Console.WriteLine("Application Deletion execption: " + graphException.ToString());
                }
            }

            //*********************************************************************************************
            // Show Batching with 3 operators.  Note: up to 5 operations can be in a batch
            //*********************************************************************************************
            // get users 
            Console.WriteLine("\n Executing Batch Request");
            BatchRequestItem firstItem = new BatchRequestItem(
                                        "GET",
                                        false,
                                        Utils.GetListUri<User>(graphConnection, null, new FilterGenerator()),
                                        null,
                                        String.Empty);

            // get members of a Group
            Uri membersUri = Utils.GetRequestUri<Group>(graphConnection, retrievedGroup.ObjectId, "members");

            BatchRequestItem secondItem = new BatchRequestItem(
                                        "GET",
                                        false,
                                        new Uri(membersUri.ToString()),
                                        null,
                                        String.Empty);

            // update an existing group's Description property
           
            retrievedGroup.Description = "New Employees in Washington State";
            
            BatchRequestItem thirdItem = new BatchRequestItem(
                                           "Patch", 
                                            true,
                                            Utils.GetRequestUri<Group>(graphConnection,retrievedGroup.ObjectId),
                                            null,
                                            retrievedGroup.ToJson(true));
            
            // Execute the batch requst
            IList<BatchRequestItem> batchRequest = new BatchRequestItem[] { firstItem, secondItem, thirdItem };
            IList<BatchResponseItem> batchResponses = graphConnection.ExecuteBatch(batchRequest);
            
            int responseCount = 0;
            foreach (BatchResponseItem responseItem in batchResponses)
            {       
                if (responseItem.Failed)
                {
                    Console.WriteLine("Failed: {0} {1}",
                                    responseItem.Exception.Code,
                                    responseItem.Exception.ErrorMessage);
                }    
                else
                {
                    Console.WriteLine("Batch Item Result {0} succeeded {1}",
                                     responseCount++,
                                     !responseItem.Failed);
                }
            }

            // this next section shows how to access the signed-in user's mailbox.
            // First we get a new token for Office365 Exchange Online Resource
            // using the multi-resource refresh token tha was included when the previoius
            // token was acquired. 
            // We can now request a new token for Office365 Exchange Online. 
            //
            string office365Emailresource = "https://outlook.office365.com/";
            string office365Token = null;
            if (userAuthnResult.IsMultipleResourceRefreshToken)
            {
                userAuthnResult = authenticationContext.AcquireTokenByRefreshToken(userAuthnResult.RefreshToken, clientIdForUserAuthn, office365Emailresource);
                office365Token = userAuthnResult.AccessToken;

                //
                // Call the Office365 API and retrieve the top 5 items from the user's mailbox.
                //
                string requestUrl = "https://outlook.office365.com/EWS/OData/Me/Inbox/Messages?$top=5";
                WebRequest getMailboxRequest;
                getMailboxRequest = WebRequest.Create(requestUrl);
                getMailboxRequest.Headers.Add(HttpRequestHeader.Authorization, "Bearer " + office365Token);
                Console.WriteLine("\n Getting the User's Mailbox Contents \n");

                //
                // Read the contents of the user's mailbox, and display to the console.
                //
                Stream objStream;
                objStream = getMailboxRequest.GetResponse().GetResponseStream();
                StreamReader objReader = new StreamReader(objStream);

                string sLine = "";
                int i = 0;

                while (sLine != null)
                {
                    i++;
                    sLine = objReader.ReadLine();
                    if (sLine != null)
                        Console.WriteLine("{0}:{1}", i, sLine);
                }
            }

            //*********************************************************************************************
            // End of Demo Console App
            //*********************************************************************************************
            Console.WriteLine("\nCompleted at {0} \n ClientRequestId: {1}", CurrentDateTime, ClientRequestId);
            Console.ReadKey();
            return;
        }
    }
}
