#region

using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Azure.ActiveDirectory.GraphClient;
using Microsoft.Azure.ActiveDirectory.GraphClient.Extensions;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

#endregion

namespace GraphConsoleAppV3
{
    public class Program
    {
        // Single-Threaded Apartment required for OAuth2 Authz Code flow (User Authn) to execute for this demo app
        [STAThread]
        private static void Main()
        {
            // record start DateTime of execution
            string currentDateTime = DateTime.Now.ToUniversalTime().ToString();

            #region Setup Active Directory Client

            //*********************************************************************
            // setup Active Directory Client
            //*********************************************************************
            ActiveDirectoryClient activeDirectoryClient;
            try
            {
                activeDirectoryClient = AuthenticationHelper.GetActiveDirectoryClientAsApplication();
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

            #endregion

            #region TenantDetails

            //*********************************************************************
            // Get Tenant Details
            // Note: update the string TenantId with your TenantId.
            // This can be retrieved from the login Federation Metadata end point:             
            // https://login.windows.net/GraphDir1.onmicrosoft.com/FederationMetadata/2007-06/FederationMetadata.xml
            //  Replace "GraphDir1.onMicrosoft.com" with any domain owned by your organization
            // The returned value from the first xml node "EntityDescriptor", will have a STS URL
            // containing your TenantId e.g. "https://sts.windows.net/4fd2b2f2-ea27-4fe5-a8f3-7b1a7c975f34/" is returned for GraphDir1.onMicrosoft.com
            //*********************************************************************
            VerifiedDomain initialDomain = new VerifiedDomain();
            VerifiedDomain defaultDomain = new VerifiedDomain();
            ITenantDetail tenant = null;
            Console.WriteLine("\n Retrieving Tenant Details");
            try
            {
                List<ITenantDetail> tenantsList = activeDirectoryClient.TenantDetails
                    .Where(tenantDetail => tenantDetail.ObjectId.Equals(Constants.TenantId))
                    .ExecuteAsync().Result.CurrentPage.ToList();
                if (tenantsList.Count > 0)
                {
                    tenant = tenantsList.First();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting TenantDetails {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }

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
                defaultDomain = tenantDetail.VerifiedDomains.First(x => x.@default.HasValue && x.@default.Value);
                Console.WriteLine("Default Domain Name: " + defaultDomain.Name);

                // Get Tenant's Tech Contacts
                foreach (string techContact in tenantDetail.TechnicalNotificationMails)
                {
                    Console.WriteLine("Tenant Tech Contact: " + techContact);
                }
            }

            #endregion

            #region Create a new User

            IUser newUser = new User();
            if (defaultDomain.Name != null)
            {
                newUser.DisplayName = "Sample App Demo User (Manager)";
                newUser.UserPrincipalName = Helper.GetRandomString(10) + "@" + defaultDomain.Name;
                newUser.AccountEnabled = true;
                newUser.MailNickname = "SampleAppDemoUserManager";
                newUser.PasswordProfile = new PasswordProfile
                {
                    Password = "TempP@ssw0rd!",
                    ForceChangePasswordNextLogin = true
                };
                newUser.UsageLocation = "US";
                try
                {
                    activeDirectoryClient.Users.AddUserAsync(newUser).Wait();
                    Console.WriteLine("\nNew User {0} was created", newUser.DisplayName);
                }
                catch (Exception e)
                {
                    Console.WriteLine("\nError creating new user {0} {1}", e.Message,
                        e.InnerException != null ? e.InnerException.Message : "");
                }
            }

            #endregion

            #region List of max 4 Users by UPN

            //*********************************************************************
            // Demonstrate Getting a list of Users with paging (get 4 users), sorted by displayName
            //*********************************************************************
            int maxUsers = 4;
            try
            {
                Console.WriteLine("\n Retrieving Users");
                List<IUser> users = activeDirectoryClient.Users.OrderBy(user =>
                    user.UserPrincipalName).Take(maxUsers).ExecuteAsync().Result.CurrentPage.ToList();
                foreach (IUser user in users)
                {
                    Console.WriteLine("UserObjectId: {0}  UPN: {1}", user.ObjectId, user.UserPrincipalName);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting Users. {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }

            #endregion

            #region Create a User with a temp Password

            //*********************************************************************************************
            // Create a new User with a temp Password
            //*********************************************************************************************
            IUser userToBeAdded = new User();
            userToBeAdded.DisplayName = "Sample App Demo User";
            userToBeAdded.UserPrincipalName = Helper.GetRandomString(10) + "@" + defaultDomain.Name;
            userToBeAdded.AccountEnabled = true;
            userToBeAdded.MailNickname = "SampleAppDemoUser";
            userToBeAdded.PasswordProfile = new PasswordProfile
            {
                Password = "TempP@ssw0rd!",
                ForceChangePasswordNextLogin = true
            };
            userToBeAdded.UsageLocation = "US";
            try
            {
                activeDirectoryClient.Users.AddUserAsync(userToBeAdded).Wait();
                Console.WriteLine("\nNew User {0} was created", userToBeAdded.DisplayName);
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError creating new user. {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }

            #endregion

            #region Update newly created User

            //*******************************************************************************************
            // update the newly created user's Password, PasswordPolicies and City
            //*********************************************************************************************
            if (userToBeAdded.ObjectId != null)
            {
                // update User's city and reset their User's Password
                userToBeAdded.City = "Seattle";
                userToBeAdded.Country = "UK";
                PasswordProfile PasswordProfile = new PasswordProfile
                {
                    Password = "newP@ssw0rd!",
                    ForceChangePasswordNextLogin = false
                };
                userToBeAdded.PasswordProfile = PasswordProfile;
                userToBeAdded.PasswordPolicies = "DisablePasswordExpiration, DisableStrongPassword";
                try
                {
                    userToBeAdded.UpdateAsync().Wait();
                    Console.WriteLine("\nUser {0} was updated", userToBeAdded.DisplayName);
                }
                catch (Exception e)
                {
                    Console.WriteLine("\nError Updating the user {0} {1}", e.Message,
                        e.InnerException != null ? e.InnerException.Message : "");
                }
            }

            #endregion

            #region Search User by UPN

            // search for a single user by UPN
            string searchString = "admin@" + initialDomain.Name;
            Console.WriteLine("\n Retrieving user with UPN {0}", searchString);
            User retrievedUser = new User();
            List<IUser> retrievedUsers = null;
            try
            {
                retrievedUsers = activeDirectoryClient.Users
                    .Where(user => user.UserPrincipalName.Equals(searchString))
                    .ExecuteAsync().Result.CurrentPage.ToList();
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting new user {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }
            // should only find one user with the specified UPN
            if (retrievedUsers != null && retrievedUsers.Count == 1)
            {
                retrievedUser = (User)retrievedUsers.First();
            }
            else
            {
                Console.WriteLine("User not found {0}", searchString);
            }

            #endregion

            #region User Operations

            if (retrievedUser.UserPrincipalName != null)
            {
                Console.WriteLine("\n Found User: " + retrievedUser.DisplayName + "  UPN: " +
                                  retrievedUser.UserPrincipalName);

                #region Assign User a Manager

                //Assigning User a new manager.
                if (newUser.ObjectId != null)
                {
                    Console.WriteLine("\n Assign User {0}, {1} as Manager.", retrievedUser.DisplayName,
                        newUser.DisplayName);
                    retrievedUser.Manager = newUser as DirectoryObject;
                    try
                    {
                        newUser.UpdateAsync().Wait();
                        Console.Write("User {0} is successfully assigned {1} as Manager.", retrievedUser.DisplayName,
                            newUser.DisplayName);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("\nError assigning manager to user. {0} {1}", e.Message,
                            e.InnerException != null ? e.InnerException.Message : "");
                    }
                }

                #endregion

                #region Get User's Manager

                //Get the retrieved user's manager.
                Console.WriteLine("\n Retrieving User {0}'s Manager.", retrievedUser.DisplayName);
                DirectoryObject usersManager = retrievedUser.Manager;
                if (usersManager != null)
                {
                    User manager = usersManager as User;
                    if (manager != null)
                    {
                        Console.WriteLine("User {0} Manager details: \nManager: {1}  UPN: {2}",
                            retrievedUser.DisplayName, manager.DisplayName, manager.UserPrincipalName);
                    }
                }
                else
                {
                    Console.WriteLine("Manager not found.");
                }

                #endregion

                #region Get User's Direct Reports

                //*********************************************************************
                // get the user's Direct Reports
                //*********************************************************************
                if (newUser.ObjectId != null)
                {
                    Console.WriteLine("\n Getting User{0}'s Direct Reports.", newUser.DisplayName);
                    IUserFetcher newUserFetcher = (IUserFetcher)newUser;
                    try
                    {
                        IPagedCollection<IDirectoryObject> directReports =
                            newUserFetcher.DirectReports.ExecuteAsync().Result;
                        do
                        {
                            List<IDirectoryObject> directoryObjects = directReports.CurrentPage.ToList();
                            foreach (IDirectoryObject directoryObject in directoryObjects)
                            {
                                if (directoryObject is User)
                                {
                                    User directReport = directoryObject as User;
                                    Console.WriteLine("User {0} Direct Report is {1}", newUser.UserPrincipalName,
                                        directReport.UserPrincipalName);
                                }
                            }
                            directReports = directReports.GetNextPageAsync().Result;
                        } while (directReports != null);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("\nError getting direct reports of user. {0} {1}", e.Message,
                            e.InnerException != null ? e.InnerException.Message : "");
                    }
                }

                #endregion

                #region Get list of Group IDS, user is member of

                //*********************************************************************
                // get a list of Group IDs that the user is a member of
                //*********************************************************************
                //const bool securityEnabledOnly = false;
                //IEnumerable<string> memberGroups = retrievedUser.GetMemberGroupsAsync(securityEnabledOnly).Result;
                //Console.WriteLine("\n {0} is a member of the following Groups (IDs)", retrievedUser.DisplayName);
                //foreach (String memberGroup in memberGroups)
                //{
                //    Console.WriteLine("Member of Group ID: " + memberGroup);
                //}

                #endregion

                #region Get User's Group And Role Membership, Getting the complete set of objects

                //*********************************************************************
                // get the User's Group and Role membership, getting the complete set of objects
                //*********************************************************************
                Console.WriteLine("\n {0} is a member of the following Group and Roles (IDs)", retrievedUser.DisplayName);
                IUserFetcher retrievedUserFetcher = retrievedUser;
                try
                {
                    IPagedCollection<IDirectoryObject> pagedCollection = retrievedUserFetcher.MemberOf.ExecuteAsync().Result;
                    do
                    {
                        List<IDirectoryObject> directoryObjects = pagedCollection.CurrentPage.ToList();
                        foreach (IDirectoryObject directoryObject in directoryObjects)
                        {
                            if (directoryObject is Group)
                            {
                                Group group = directoryObject as Group;
                                Console.WriteLine(" Group: {0}  Description: {1}", group.DisplayName, group.Description);
                            }
                            if (directoryObject is DirectoryRole)
                            {
                                DirectoryRole role = directoryObject as DirectoryRole;
                                Console.WriteLine(" Role: {0}  Description: {1}", role.DisplayName, role.Description);
                            }
                        }
                        pagedCollection = pagedCollection.GetNextPageAsync().Result;
                    } while (pagedCollection != null);
                }
                catch (Exception e)
                {
                    Console.WriteLine("\nError getting user's groups and roles memberships. {0} {1}", e.Message,
                        e.InnerException != null ? e.InnerException.Message : "");
                }

                #endregion
            }

            #endregion

            #region Search for User (People Picker)

            //*********************************************************************
            // People picker
            // Search for a user using text string "Us" match against userPrincipalName, displayName, giveName, surname   
            //*********************************************************************
            searchString = "Us";
            Console.WriteLine("\nSearching for any user with string {0} in UPN,DisplayName,First or Last Name",
                searchString);
            List<IUser> usersList = null;
            IPagedCollection<IUser> searchResults = null;
            try
            {
                IUserCollection userCollection = activeDirectoryClient.Users;
                searchResults = userCollection.Where(user =>
                    user.UserPrincipalName.StartsWith(searchString) ||
                    user.DisplayName.StartsWith(searchString) ||
                    user.GivenName.StartsWith(searchString) ||
                    user.Surname.StartsWith(searchString)).ExecuteAsync().Result;
                usersList = searchResults.CurrentPage.ToList();
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting User {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }

            if (usersList != null && usersList.Count > 0)
            {
                do
                {
                    usersList = searchResults.CurrentPage.ToList();
                    foreach (IUser user in usersList)
                    {
                        Console.WriteLine("User DisplayName: {0}  UPN: {1}",
                            user.DisplayName, user.UserPrincipalName);
                    }
                    searchResults = searchResults.GetNextPageAsync().Result;
                } while (searchResults != null);
            }
            else
            {
                Console.WriteLine("User not found");
            }

            #endregion

            #region Search for Group using StartWith filter

            //*********************************************************************
            // Search for a group using a startsWith filter (displayName property)
            //*********************************************************************
            Group retrievedGroup = new Group();
            searchString = "My";
            List<IGroup> foundGroups = null;
            try
            {
                foundGroups = activeDirectoryClient.Groups
                    .Where(group => group.DisplayName.StartsWith(searchString))
                    .ExecuteAsync().Result.CurrentPage.ToList();
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting Group {0} {1}",
                    e.Message, e.InnerException != null ? e.InnerException.Message : "");
            }
            if (foundGroups != null && foundGroups.Count > 0)
            {
                retrievedGroup = foundGroups.First() as Group;
            }
            else
            {
                Console.WriteLine("Group Not Found");
            }

            #endregion

            #region Assign Member to Group

            if (retrievedGroup.ObjectId != null)
            {
                try
                {
                    retrievedGroup.Members.Add(newUser as DirectoryObject);
                    retrievedGroup.UpdateAsync().Wait();
                }
                catch (Exception e)
                {
                    Console.WriteLine("\nError assigning member to group. {0} {1}",
                        e.Message, e.InnerException != null ? e.InnerException.Message : "");
                }

            }

            #endregion

            #region Get Group members

            if (retrievedGroup.ObjectId != null)
            {
                Console.WriteLine("\n Found Group: " + retrievedGroup.DisplayName + "  " + retrievedGroup.Description);

                //*********************************************************************
                // get the groups' membership - 
                // Note this method retrieves ALL links in one request - please use this method with care - this
                // may return a very large number of objects
                //*********************************************************************
                IGroupFetcher retrievedGroupFetcher = retrievedGroup;
                try
                {
                    IPagedCollection<IDirectoryObject> members = retrievedGroupFetcher.Members.ExecuteAsync().Result;
                    Console.WriteLine(" Members:");
                    do
                    {
                        List<IDirectoryObject> directoryObjects = members.CurrentPage.ToList();
                        foreach (IDirectoryObject member in directoryObjects)
                        {
                            if (member is User)
                            {
                                User user = member as User;
                                Console.WriteLine("User DisplayName: {0}  UPN: {1}",
                                    user.DisplayName,
                                    user.UserPrincipalName);
                            }
                            if (member is Group)
                            {
                                Group group = member as Group;
                                Console.WriteLine("Group DisplayName: {0}", group.DisplayName);
                            }
                            if (member is Contact)
                            {
                                Contact contact = member as Contact;
                                Console.WriteLine("Contact DisplayName: {0}", contact.DisplayName);
                            }
                        }
                        members = members.GetNextPageAsync().Result;
                    } while (members != null);
                }
                catch (Exception e)
                {
                    Console.WriteLine("\nError getting groups' membership. {0} {1}",
                        e.Message, e.InnerException != null ? e.InnerException.Message : "");
                }
            }

            #endregion

            #region Add User to Group

            //*********************************************************************************************
            // Add User to the "WA" Group 
            //*********************************************************************************************
            if (retrievedGroup.ObjectId != null)
            {
                try
                {
                    retrievedGroup.Members.Add(userToBeAdded as DirectoryObject);
                    retrievedGroup.UpdateAsync().Wait();
                }
                catch (Exception e)
                {
                    Console.WriteLine("\nAdding user to group failed {0} {1}", e.Message,
                        e.InnerException != null ? e.InnerException.Message : "");
                }
            }

            #endregion

            #region Create a new Group

            //*********************************************************************************************
            // Create a new Group
            //*********************************************************************************************
            Group californiaEmployees = new Group
            {
                DisplayName = "California Employees" + Helper.GetRandomString(8),
                Description = "Employees in the state of California",
                MailNickname = "CalEmployees",
                MailEnabled = false,
                SecurityEnabled = true
            };
            try
            {
                activeDirectoryClient.Groups.AddGroupAsync(californiaEmployees).Wait();
                Console.WriteLine("\nNew Group {0} was created", californiaEmployees.DisplayName);
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError creating new Group {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }

            #endregion

            #region Delete User

            //*********************************************************************************************
            // Delete the user that we just created
            //*********************************************************************************************
            if (userToBeAdded.ObjectId != null)
            {
                try
                {
                    userToBeAdded.DeleteAsync().Wait();
                    Console.WriteLine("\nUser {0} was deleted", userToBeAdded.DisplayName);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Deleting User failed {0} {1}", e.Message,
                        e.InnerException != null ? e.InnerException.Message : "");
                }
            }
            if (newUser.ObjectId != null)
            {
                try
                {
                    newUser.DeleteAsync().Wait();
                    Console.WriteLine("\nUser {0} was deleted", newUser.DisplayName);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Deleting User failed {0} {1}", e.Message,
                        e.InnerException != null ? e.InnerException.Message : "");
                }
            }

            #endregion

            #region Delete Group

            //*********************************************************************************************
            // Delete the Group that we just created
            //*********************************************************************************************
            if (californiaEmployees.ObjectId != null)
            {
                try
                {
                    californiaEmployees.DeleteAsync().Wait();
                    Console.WriteLine("\nGroup {0} was deleted", californiaEmployees.DisplayName);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Deleting Group failed {0} {1}", e.Message,
                        e.InnerException != null ? e.InnerException.Message : "");
                }
            }

            #endregion

            #region Get All Roles

            //*********************************************************************
            // Get All Roles
            //*********************************************************************
            List<IDirectoryRole> foundRoles = null;
            try
            {
                foundRoles = activeDirectoryClient.DirectoryRoles.ExecuteAsync().Result.CurrentPage.ToList();
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting Roles {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }

            if (foundRoles != null && foundRoles.Count > 0)
            {
                foreach (IDirectoryRole role in foundRoles)
                {
                    Console.WriteLine("\n Found Role: {0} {1} {2} ",
                        role.DisplayName, role.Description, role.ObjectId);
                }
            }
            else
            {
                Console.WriteLine("Role Not Found {0}", searchString);
            }

            #endregion

            #region Get Service Principals

            //*********************************************************************
            // get the Service Principals
            //*********************************************************************
            IPagedCollection<IServicePrincipal> servicePrincipals = null;
            try
            {
                servicePrincipals = activeDirectoryClient.ServicePrincipals.ExecuteAsync().Result;
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting Service Principal {0} {1}",
                    e.Message, e.InnerException != null ? e.InnerException.Message : "");
            }
            if (servicePrincipals != null)
            {
                do
                {
                    List<IServicePrincipal> servicePrincipalsList = servicePrincipals.CurrentPage.ToList();
                    foreach (IServicePrincipal servicePrincipal in servicePrincipalsList)
                    {
                        Console.WriteLine("Service Principal AppId: {0}  Name: {1}", servicePrincipal.AppId,
                            servicePrincipal.DisplayName);
                    }
                    servicePrincipals = servicePrincipals.GetNextPageAsync().Result;
                } while (servicePrincipals != null);
            }

            #endregion

            #region Get Applications

            //*********************************************************************
            // get the Application objects
            //*********************************************************************
            IPagedCollection<IApplication> applications = null;
            try
            {
                applications = activeDirectoryClient.Applications.Take(999).ExecuteAsync().Result;
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting Applications {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }
            if (applications != null)
            {
                do
                {
                    List<IApplication> appsList = applications.CurrentPage.ToList();
                    foreach (IApplication app in appsList)
                    {
                        Console.WriteLine("Application AppId: {0}  Name: {1}", app.AppId, app.DisplayName);
                    }
                    applications = applications.GetNextPageAsync().Result;
                } while (applications != null);
            }

            #endregion

            #region User License Assignment

            //*********************************************************************************************
            // User License Assignment - assign EnterprisePack license to new user, and disable SharePoint service
            //   first get a list of Tenant's subscriptions and find the "Enterprisepack" one
            //   Enterprise Pack includes service Plans for ExchangeOnline, SharePointOnline and LyncOnline
            //   validate that Subscription is Enabled and there are enough units left to assign to users
            //*********************************************************************************************
            IPagedCollection<ISubscribedSku> skus = null;
            try
            {
                skus = activeDirectoryClient.SubscribedSkus.ExecuteAsync().Result;
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting Applications {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }
            if (skus != null)
            {
                do
                {
                    List<ISubscribedSku> subscribedSkus = skus.CurrentPage.ToList();
                    foreach (ISubscribedSku sku in subscribedSkus)
                    {
                        if (sku.SkuPartNumber == "ENTERPRISEPACK")
                        {
                            if ((sku.PrepaidUnits.Enabled.Value > sku.ConsumedUnits) &&
                                (sku.CapabilityStatus == "Enabled"))
                            {
                                // create addLicense object and assign the Enterprise Sku GUID to the skuId
                                // 
                                AssignedLicense addLicense = new AssignedLicense { SkuId = sku.SkuId.Value };

                                // find plan id of SharePoint Service Plan
                                foreach (ServicePlanInfo servicePlan in sku.ServicePlans)
                                {
                                    if (servicePlan.ServicePlanName.Contains("SHAREPOINT"))
                                    {
                                        addLicense.DisabledPlans.Add(servicePlan.ServicePlanId.Value);
                                        break;
                                    }
                                }

                                IList<AssignedLicense> licensesToAdd = new[] { addLicense };
                                IList<Guid> licensesToRemove = new Guid[] { };

                                // attempt to assign the license object to the new user 
                                try
                                {
                                    if (newUser.ObjectId != null)
                                    {
                                        newUser.AssignLicenseAsync(licensesToAdd, licensesToRemove).Wait();
                                        Console.WriteLine("\n User {0} was assigned license {1}",
                                            newUser.DisplayName,
                                            addLicense.SkuId);
                                    }
                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine("\nLicense assingment failed {0} {1}", e.Message,
                                        e.InnerException != null ? e.InnerException.Message : "");
                                }
                            }
                        }
                    }
                    skus = skus.GetNextPageAsync().Result;
                } while (skus != null);
            }

            #endregion

            #region Switch to OAuth Authorization Code Grant (Acting as a user)

            activeDirectoryClient = AuthenticationHelper.GetActiveDirectoryClientAsUser();

            #endregion

            #region Create Application

            //*********************************************************************************************
            // Create a new Application object with App Role Assignment (Direct permission)
            //*********************************************************************************************
            Application appObject = new Application { DisplayName = "Test-Demo App" + Helper.GetRandomString(8) };
            appObject.IdentifierUris.Add("https://localhost/demo/" + Guid.NewGuid());
            appObject.ReplyUrls.Add("https://localhost/demo");
            AppRole appRole = new AppRole();
            appRole.Id = Guid.NewGuid();
            appRole.IsEnabled = true;
            appRole.AllowedMemberTypes.Add("User");
            appRole.DisplayName = "Something";
            appRole.Description = "Anything";
            appRole.Value = "policy.write";
            appObject.AppRoles.Add(appRole);

            // created Keycredential object for the new App object
            KeyCredential keyCredential = new KeyCredential
            {
                StartDate = DateTime.UtcNow,
                EndDate = DateTime.UtcNow.AddYears(1),
                Type = "Symmetric",
                Value = Convert.FromBase64String("g/TMLuxgzurjQ0Sal9wFEzpaX/sI0vBP3IBUE/H/NS4="),
                Usage = "Verify"
            };
            appObject.KeyCredentials.Add(keyCredential);

            try
            {
                activeDirectoryClient.Applications.AddApplicationAsync(appObject).Wait();
                Console.WriteLine("New Application created: " + appObject.ObjectId);
            }
            catch (Exception e)
            {
                Console.WriteLine("Application Creation execption: {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }

            #endregion

            #region Create Service Principal

            //*********************************************************************************************
            // create a new Service principal
            //*********************************************************************************************
            ServicePrincipal newServicePrincpal = new ServicePrincipal();
            if (appObject != null)
            {
                newServicePrincpal.DisplayName = appObject.DisplayName;
                newServicePrincpal.AccountEnabled = true;
                newServicePrincpal.AppId = appObject.AppId;
                try
                {
                    activeDirectoryClient.ServicePrincipals.AddServicePrincipalAsync(newServicePrincpal).Wait();
                    Console.WriteLine("New Service Principal created: " + newServicePrincpal.ObjectId);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Service Principal Creation execption: {0} {1}", e.Message,
                        e.InnerException != null ? e.InnerException.Message : "");
                }
            }

            #endregion

            #region Create an Extension Property

            ExtensionProperty linkedInUserId = new ExtensionProperty
            {
                Name = "linkedInUserId",
                DataType = "String",
                TargetObjects = { "User" }
            };
            try
            {
                appObject.ExtensionProperties.Add(linkedInUserId);
                appObject.UpdateAsync().Wait();
                Console.WriteLine("\nUser object extended successfully.");
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError extending the user object {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }

            #endregion

            #region Manipulate an Extension Property

            try
            {
                if (retrievedUser != null && retrievedUser.ObjectId != null)
                {
                    retrievedUser.SetExtendedProperty(linkedInUserId.Name, "ExtensionPropertyValue");
                    retrievedUser.UpdateAsync().Wait();
                    Console.WriteLine("\nUser {0}'s extended property set successully.", retrievedUser.DisplayName);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError Updating the user object {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }

            #endregion

            #region Get an Extension Property

            try
            {
                if (retrievedUser != null && retrievedUser.ObjectId != null)
                {
                    IReadOnlyDictionary<string, object> extendedProperties = retrievedUser.GetExtendedProperties();
                    object extendedProperty = extendedProperties[linkedInUserId.Name];
                    Console.WriteLine("\n Retrieved User {0}'s extended property value is: {1}.", retrievedUser.DisplayName,
                        extendedProperty);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError Updating the user object {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }

            #endregion

            #region Assign Direct Permission
            try
            {
                User user =
                        (User)activeDirectoryClient.Users.ExecuteAsync().Result.CurrentPage.ToList().FirstOrDefault();
                if (appObject.ObjectId != null && user != null && newServicePrincpal.ObjectId != null)
                {
                    AppRoleAssignment appRoleAssignment = new AppRoleAssignment();
                    appRoleAssignment.Id = appRole.Id;
                    appRoleAssignment.ResourceId = Guid.Parse(newServicePrincpal.ObjectId);
                    appRoleAssignment.PrincipalType = "User";
                    appRoleAssignment.PrincipalId = Guid.Parse(user.ObjectId);
                    user.AppRoleAssignments.Add(appRoleAssignment);
                    user.UpdateAsync().Wait();
                    Console.WriteLine("User {0} is successfully assigned direct permission.", retrievedUser.DisplayName);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Direct Permission Assignment failed: {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }

            #endregion

            #region Get Devices

            //*********************************************************************************************
            // Get a list of Mobile Devices from tenant
            //*********************************************************************************************
            Console.WriteLine("\nGetting Devices");
            IPagedCollection<IDevice> devices = null;
            try
            {
                devices = activeDirectoryClient.Devices.ExecuteAsync().Result;
            }
            catch (Exception e)
            {
                Console.WriteLine("/nError getting devices {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }

            if (devices != null)
            {
                do
                {
                    List<IDevice> devicesList = devices.CurrentPage.ToList();
                    foreach (IDevice device in devicesList)
                    {
                        if (device.ObjectId != null)
                        {
                            Console.WriteLine("Device ID: {0}, Type: {1}", device.DeviceId, device.DeviceOSType);
                            IPagedCollection<IDirectoryObject> registeredOwners = device.RegisteredOwners;
                            if (registeredOwners != null)
                            {
                                do
                                {
                                    List<IDirectoryObject> registeredOwnersList = registeredOwners.CurrentPage.ToList();
                                    foreach (IDirectoryObject owner in registeredOwnersList)
                                    {
                                        Console.WriteLine("Device Owner ID: " + owner.ObjectId);
                                    }
                                    registeredOwners = registeredOwners.GetNextPageAsync().Result;
                                } while (registeredOwners != null);
                            }
                        }
                    }
                    devices = devices.GetNextPageAsync().Result;
                } while (devices != null);
            }

            #endregion

            #region Create New Permission

            //*********************************************************************************************
            // Create new permission object
            //*********************************************************************************************
            OAuth2PermissionGrant permissionObject = new OAuth2PermissionGrant();
            permissionObject.ConsentType = "AllPrincipals";
            permissionObject.Scope = "user_impersonation";
            permissionObject.StartTime = DateTime.Now;
            permissionObject.ExpiryTime = (DateTime.Now).AddMonths(12);

            // resourceId is objectId of the resource, in this case objectId of AzureAd (Graph API)
            permissionObject.ResourceId = "52620afb-80de-4096-a826-95f4ad481686";

            //ClientId = objectId of servicePrincipal
            permissionObject.ClientId = newServicePrincpal.ObjectId;
            try
            {
                activeDirectoryClient.Oauth2PermissionGrants.AddOAuth2PermissionGrantAsync(permissionObject).Wait();
                Console.WriteLine("New Permission object created: " + permissionObject.ObjectId);
            }
            catch (Exception e)
            {
                Console.WriteLine("Permission Creation exception: {0} {1}", e.Message, e.InnerException != null ? e.InnerException.Message : "");
            }

            #endregion

            #region Get All Permissions

            //*********************************************************************************************
            // get all Permission Objects
            //*********************************************************************************************
            Console.WriteLine("\n Getting Permissions");
            IPagedCollection<IOAuth2PermissionGrant> permissions = null;
            try
            {
                permissions = activeDirectoryClient.Oauth2PermissionGrants.ExecuteAsync().Result;
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: {0} {1}", e.Message, e.InnerException != null ? e.InnerException.Message : "");
            }
            if (permissions != null)
            {
                do
                {
                    List<IOAuth2PermissionGrant> perms = permissions.CurrentPage.ToList();
                    foreach (IOAuth2PermissionGrant perm in perms)
                    {
                        Console.WriteLine("Permission: {0}  Name: {1}", perm.ClientId, perm.Scope);
                    }
                    permissions = permissions.GetNextPageAsync().Result;
                } while (permissions != null);
            }

            #endregion

            #region Delete Application

            //*********************************************************************************************
            // Delete Application Objects
            //*********************************************************************************************
            if (appObject.ObjectId != null)
            {
                try
                {
                    appObject.DeleteAsync().Wait();
                    Console.WriteLine("Deleted Application object: " + appObject.ObjectId);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Application Deletion execption: {0} {1}", e.Message,
                        e.InnerException != null ? e.InnerException.Message : "");
                }
            }

            #endregion

            #region Batch Operations

            //*********************************************************************************************
            // Show Batching with 3 operators.  Note: up to 5 operations can be in a batch
            //*********************************************************************************************
            IReadOnlyQueryableSet<User> userQuery = activeDirectoryClient.DirectoryObjects.OfType<User>();
            IReadOnlyQueryableSet<Group> groupsQuery = activeDirectoryClient.DirectoryObjects.OfType<Group>();
            IReadOnlyQueryableSet<DirectoryRole> rolesQuery =
                activeDirectoryClient.DirectoryObjects.OfType<DirectoryRole>();
            try
            {
                IBatchElementResult[] batchResult =
                    activeDirectoryClient.Context.ExecuteBatchAsync(userQuery, groupsQuery, rolesQuery).Result;
                int responseCount = 1;
                foreach (IBatchElementResult result in batchResult)
                {
                    if (result.FailureResult != null)
                    {
                        Console.WriteLine("Failed: {0} ",
                            result.FailureResult.InnerException);
                    }
                    if (result.SuccessResult != null)
                    {
                        Console.WriteLine("Batch Item Result {0} succeeded",
                            responseCount++);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Batch execution failed. : {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }

            #endregion

            //*********************************************************************************************
            // End of Demo Console App
            //*********************************************************************************************

            Console.WriteLine("\nCompleted at {0} \n Press Any Key to Exit.", currentDateTime);
            Console.ReadKey();
        }
    }
}