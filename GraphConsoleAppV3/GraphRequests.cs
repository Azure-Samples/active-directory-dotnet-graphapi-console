#region

using Microsoft.Azure.ActiveDirectory.GraphClient;
using Microsoft.Azure.ActiveDirectory.GraphClient.Extensions;
using System;
using System.Collections.Generic;
using System.Data.Services.Client;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using Application = Microsoft.Azure.ActiveDirectory.GraphClient.Application;

#endregion

namespace GraphConsoleAppV3
{
    internal class Requests
    {
        public static async Task UserMode()
        {

            ActiveDirectoryClient client;

            //*********************************************************************
            // setup Microsoft Graph client for user...
            //*********************************************************************
            try
            {
                client = AuthenticationHelper.GetActiveDirectoryClientAsUser();
            }
            catch (Exception e)
            {
                Program.WriteError("Acquiring a token failed with the following error: {0}",
                    Program.ExtractErrorMessage(e));
                //TODO: Implement retry and back-off logic per the guidance given here:http://msdn.microsoft.com/en-us/library/dn168916.aspx
                Console.ReadKey();
                return;
            }


            User newUser = null;
            Application newApp = null;
            IDomain newDomain = null;
            Group newGroup = null;
            ExtensionProperty newExtension = null;
            ServicePrincipal newServicePrincipal = null;
            OAuth2PermissionGrant newPermissionGrant = null;

            try
            {
                Console.WriteLine("\nStarting user-mode requests...");
                Console.WriteLine("\n=============================\n\n");


                ITenantDetail tenantDetail = await GetTenantDetails(client, UserModeConstants.TenantId);
                User signedInUser = await GetSignedInUser(client);
                await UpdateUsersPhoto(signedInUser);
                await PrintUsersGroupsAndRoles(signedInUser);

                Console.WriteLine("\nSearching for any user based on UPN, DisplayName, First or Last Name");
                Console.WriteLine("\nPlease enter the user's name you are looking for:");
                string searchString = Console.ReadLine();
                await PeoplePickerExample(client, searchString);
                await PrintUsersManager(signedInUser);

                newUser = await CreateNewUser(client,
                    tenantDetail.VerifiedDomains.First(x => x.@default.HasValue && x.@default.Value));
                await UpdateNewUser(newUser);
                await ResetUserPassword(newUser);
                await AssignManager(newUser, signedInUser);
                await AssignLicenses(client, newUser);

                await PrintGroupMembers(client);
                newGroup = await CreateNewGroup(client);
                await AddUserToGroup(newGroup, newUser);
                await RemoveUserFromGroup(newGroup, newUser);

                await PrintAllRoles(client, searchString);
                await PrintServicePrincipals(client);
                await PrintApplications(client);

                newApp = await CreateNewApplication(client, newUser);
                newServicePrincipal = await CreateServicePrincipal(client, newApp);
                newExtension = await CreateSchemaExtensions(client, newApp);
                await SetExtensionProperty(newExtension, signedInUser);
                PrintExtensionProperty(newUser, newExtension);

                await AssignAppRole(client, newApp, newServicePrincipal);
                newPermissionGrant = await CreateOAuth2Permission(client, newServicePrincipal);
                await PrintDevicesAndOwners(client);
                await PrintAllPermissions(client);

                await PrintAllDomains(client);
                newDomain = await CreateNewDomain(client);
                PrintDomainVerificationDetails(newDomain as IDomainFetcher);
                VerifyDomain(newDomain);

                await BatchOperations(client);
            }
            finally
            {
                DeleteUser(newUser).Wait();
                DeleteDomain(newDomain).Wait();
                DeleteGroup(newGroup).Wait();
                DeleteServicePrincipalAndPermission(newServicePrincipal, newPermissionGrant).Wait();
                DeleteApplication(newApp, newExtension).Wait();
            }
        }

        public static async Task AppMode()
        {

            ActiveDirectoryClient client;
            //*********************************************************************
            // setup Microsoft Graph client for app
            //*********************************************************************
            try
            {
                client = AuthenticationHelper.GetActiveDirectoryClientAsApplication();
            }
            catch (Exception e)
            {
                //TODO: Implement retry and back-off logic per the guidance given here:http://msdn.microsoft.com/en-us/library/dn168916.aspx
                Program.WriteError("Acquiring a token failed with the following error: {0}",
                    Program.ExtractErrorMessage(e));
                return;
            }


            Console.WriteLine("\nStarting app-mode requests...");
            Console.WriteLine("\nAll requests are done in the context of the application only (daemon style app)\n\n");
            Console.WriteLine("\n=============================\n\n");

            Console.WriteLine("\nSearching for any user based on UPN, DisplayName, First or Last Name");
            Console.WriteLine("\nPlease enter the user's name you are looking for:");
            string searchString = Console.ReadLine();

            await PeoplePickerExample(client, searchString);
        }

        private static async Task<ITenantDetail> GetTenantDetails(IActiveDirectoryClient client, string tenantId)
        {
            //*********************************************************************
            // The following section may be run by any user, as long as the app
            // has been granted the minimum of User.Read (and User.ReadWrite to update photo)
            // and User.ReadBasic.All scope permissions. Directory.ReadWrite.All
            // or Directory.AccessAsUser.All will also work, but are much more privileged.
            //*********************************************************************


            //*********************************************************************
            // Get Tenant Details
            // Note: update the string TenantId with your TenantId.
            // This can be retrieved from the login Federation Metadata end point:             
            // https://login.windows.net/GraphDir1.onmicrosoft.com/FederationMetadata/2007-06/FederationMetadata.xml
            //  Replace "GraphDir1.onMicrosoft.com" with any domain owned by your organization
            // The returned value from the first xml node "EntityDescriptor", will have a STS URL
            // containing your TenantId e.g. "https://sts.windows.net/4fd2b2f2-ea27-4fe5-a8f3-7b1a7c975f34/" is returned for GraphDir1.onMicrosoft.com
            //*********************************************************************

            ITenantDetail tenant = null;
            Console.WriteLine("\n Retrieving Tenant Details");

            try
            {
                IPagedCollection<ITenantDetail> tenantsCollection = await client.TenantDetails
                    .Where(tenantDetail => tenantDetail.ObjectId.Equals(tenantId))
                    .ExecuteAsync();
                List<ITenantDetail> tenantsList = tenantsCollection.CurrentPage.ToList();

                if (tenantsList.Count > 0)
                {
                    tenant = tenantsList.First();
                }
            }
            catch (Exception e)
            {
                Program.WriteError("\nError getting TenantDetails {0}", Program.ExtractErrorMessage(e));
            }

            if (tenant == null)
            {
                Console.WriteLine("Tenant not found");
                return null;
            }
            else
            {
                TenantDetail tenantDetail = (TenantDetail) tenant;
                Console.WriteLine("Tenant Display Name: " + tenantDetail.DisplayName);

                // Get the Tenant's Verified Domains 
                var initialDomain = tenantDetail.VerifiedDomains.First(x => x.Initial.HasValue && x.Initial.Value);
                Console.WriteLine("Initial Domain Name: " + initialDomain.Name);
                var defaultDomain = tenantDetail.VerifiedDomains.First(x => x.@default.HasValue && x.@default.Value);
                Console.WriteLine("Default Domain Name: " + defaultDomain.Name);

                // Get Tenant's Tech Contacts
                foreach (string techContact in tenantDetail.TechnicalNotificationMails)
                {
                    Console.WriteLine("Tenant Tech Contact: " + techContact);
                }
                return tenantDetail;
            }

        }

        #region operations with users

        private static async Task<User> GetSignedInUser(ActiveDirectoryClient client)
        {


            User signedInUser = new User();
            try
            {
                signedInUser = (User) await client.Me.ExecuteAsync();
                Console.WriteLine("\nUser UPN: {0}, DisplayName: {1}", signedInUser.UserPrincipalName,
                    signedInUser.DisplayName);
            }
            catch (Exception e)
            {
                Program.WriteError("\nError getting signed in user {0}", Program.ExtractErrorMessage(e));
            }

            if (signedInUser.ObjectId != null)
            {
                IUser sUser = (IUser) signedInUser;
                IStreamFetcher photo = (IStreamFetcher) sUser.ThumbnailPhoto;
                try
                {
                    DataServiceStreamResponse response =
                        await photo.DownloadAsync();
                    Console.WriteLine("\nUser {0} GOT thumbnailphoto", signedInUser.DisplayName);
                }
                catch (Exception e)
                {
                    Program.WriteError("\nError getting the user's photo - may not exist {0}",
                        Program.ExtractErrorMessage(e));
                }
            }
            return signedInUser;

        }

        private static async Task PeoplePickerExample(IActiveDirectoryClient client, string searchString)
        {

            //*********************************************************************
            // People picker
            // Search for a user using text string "Us" match against userPrincipalName, displayName, giveName, surname
            // Requires minimum of User.ReadBasic.All.
            //*********************************************************************

            List<IUser> usersList = null;
            IPagedCollection<IUser> searchResults = null;
            try
            {
                IUserCollection userCollection = client.Users;
                searchResults = await userCollection.Where(user =>
                    user.UserPrincipalName.StartsWith(searchString) ||
                    user.DisplayName.StartsWith(searchString) ||
                    user.GivenName.StartsWith(searchString) ||
                    user.Surname.StartsWith(searchString)).Take(10).ExecuteAsync();
                usersList = searchResults.CurrentPage.ToList();
            }
            catch (Exception e)
            {
                Program.WriteError("\nError getting User {0}", Program.ExtractErrorMessage(e));
            }

            if (usersList != null && usersList.Count > 0)
            {
                do
                {
                    int index = 1;
                    usersList = searchResults.CurrentPage.ToList();
                    foreach (IUser user in usersList)
                    {
                        Console.WriteLine("User {0} DisplayName: {1}  UPN: {2}",
                            index, user.DisplayName, user.UserPrincipalName);
                        index++;
                    }
                    searchResults = await searchResults.GetNextPageAsync();
                } while (searchResults != null);
            }
            else
            {
                Console.WriteLine("User not found");
            }

        }

        private static async Task PrintUsersManager(User user)
        {

            // ***********************************************************************
            // NOTE:  This requires User.Read.All permission scope, or Directory.Read.All or Directory.AccessAsUser.All
            // Group membership requires Group.Read.All or Directory.Read.All (the latter is required for role memberships)
            // Code snippet also demonstrates paging through user's direct reports
            // ***********************************************************************

            // manager and reports...
            try
            {
                Console.WriteLine("\nRetrieving signed in user's Manager and Direct Reports");
                IUserFetcher userFetcher = user as IUserFetcher;
                IDirectoryObject manager = await userFetcher.Manager.ExecuteAsync();
                IPagedCollection<IDirectoryObject> reports = await userFetcher.DirectReports.ExecuteAsync();

                if (manager is User)
                {
                    Console.WriteLine("\n  Manager (user):" + ((IUser) (manager)).DisplayName);
                }
                else if (manager is Contact)
                {
                    Console.WriteLine("\n  Manager (contact):" + ((IContact) (manager)).DisplayName);
                }
                else
                {
                    Console.WriteLine("\n  User has no manager :)");
                }

                if (reports != null)
                {
                    Console.WriteLine("\n  Direct reports:");

                    do
                    {
                        List<IDirectoryObject> directoryObjects = reports.CurrentPage.ToList();
                        foreach (IDirectoryObject directoryObject in directoryObjects)
                        {
                            if (directoryObject is User)
                            {
                                Console.WriteLine("\n    " + ((IUser) (manager)).DisplayName);
                            }
                            else if (directoryObject is Contact)
                            {
                                Console.WriteLine("\n    " + ((IContact) (manager)).DisplayName);
                            }

                        }
                        reports = await reports.GetNextPageAsync();
                    } while (reports != null);
                }

                else
                {
                    Console.WriteLine("\n User has no direct reports");
                }
            }
            catch (Exception e)
            {
                Program.WriteError("\nError getting user's manager and reports {0}", Program.ExtractErrorMessage(e));
            }

        }

        private static async Task PrintUsersGroupsAndRoles(User user)
        {

            Console.WriteLine("\n Signed in user {0} is a member of the following Group and Roles (IDs)",
                user.DisplayName);
            IUserFetcher signedInUserFetcher = user;
            try
            {
                IPagedCollection<IDirectoryObject> pagedCollection = await signedInUserFetcher.MemberOf.ExecuteAsync();
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
                    pagedCollection = await pagedCollection.GetNextPageAsync();
                } while (pagedCollection != null);
            }
            catch (Exception e)
            {
                Program.WriteError("\nError getting Signed in user's groups and roles memberships. {0}",
                    Program.ExtractErrorMessage(e));
            }

        }

        private static async Task<User> CreateNewUser(IActiveDirectoryClient client, VerifiedDomain defaultDomain)
        {

            // **********************************************************
            // Requires Directory.ReadWrite.All or Directory.AccessAsUser.All, and the signed in user
            // must be a privileged user (like a company or user admin)
            // **********************************************************

            User newUser = new User();
            if (defaultDomain.Name != null)
            {
                Console.WriteLine("\nCreating a new user...");
                Console.WriteLine("\n  Please enter first name for new user:");
                String firstName = Console.ReadLine();
                Console.WriteLine("\n  Please enter last name for new user:");
                String lastName = Console.ReadLine();
                newUser.DisplayName = firstName + " " + lastName;
                newUser.UserPrincipalName = firstName + "." + lastName + Helper.GetRandomString(4) + "@" +
                                            defaultDomain.Name;
                newUser.AccountEnabled = true;
                newUser.MailNickname = firstName + lastName;
                newUser.PasswordProfile = new PasswordProfile
                {
                    Password = "ChangeMe123!",
                    ForceChangePasswordNextLogin = true
                };
                newUser.UsageLocation = "US";
                try
                {
                    await client.Users.AddUserAsync(newUser);
                    Console.WriteLine("\nNew User {0} was created", newUser.DisplayName);
                }
                catch (Exception e)
                {
                    Program.WriteError("\nError creating new user {0}", Program.ExtractErrorMessage(e));
                }
            }
            return newUser;

        }

        private static async Task UpdateUsersPhoto(IUser user)
        {

            // NOTE:  Updating the signed in user's photo requires User.ReadWrite (when available) or 
            // Directory.ReadWrite.All or Directory.AccessAsUser.All
            if (user.ObjectId != null)
            {
                Console.WriteLine("\nDo you want to update your thumbnail photo? yes/no\n");
                string update = Console.ReadLine();

                if (update != null && update.Equals("yes"))
                {
                    try
                    {
                        string photo = "GraphConsoleAppV3.Resources.default.PNG";
                        Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(photo);

                        await user.ThumbnailPhoto.UploadAsync(stream, "application/image");
                        Console.WriteLine("\nUser {0} was updated with a thumbnailphoto", user.DisplayName);
                    }
                    catch (Exception e)
                    {
                        Program.WriteError("\nError Updating the user photo {0}", Program.ExtractErrorMessage(e));
                    }
                }
            }


        }

        private static async Task UpdateNewUser(IUser newUser)
        {

            //*******************************************************************************************
            // update the newly created user's Password, PasswordPolicies and City
            //*********************************************************************************************
            if (newUser.ObjectId != null)
            {
                // update User's info
                newUser.City = "Seattle";
                newUser.Country = "UK";
                newUser.Mobile = "+4477889456789";
                newUser.UserType = "Member";

                try
                {
                    await newUser.UpdateAsync();
                    Console.WriteLine("\nUser {0} was updated:", newUser.DisplayName);
                    Console.WriteLine("\t{0}:\t{1}", "City", newUser.City);
                    Console.WriteLine("\t{0}:\t{1}", "Country", newUser.Country);
                    Console.WriteLine("\t{0}:\t{1}", "Mobile", newUser.Mobile);
                    Console.WriteLine("\t{0}:\t{1}", "Type", newUser.UserType);
                }
                catch (Exception e)
                {
                    Program.WriteError("\nError Updating the user {0}", Program.ExtractErrorMessage(e));
                }
            }

        }

        private static async Task ResetUserPassword(IUser user)
        {

            //*******************************************************************************************
            // update the newly created user's Password and PasswordPolicies
            // requires Directory.AccessAsUser.All and that the current user is a user, helpdesk or company admin
            //*********************************************************************************************
            if (user.ObjectId != null)
            {
                // update User's password policy and reset password - forcing change password at next logon
                PasswordProfile PasswordProfile = new PasswordProfile
                {
                    Password = "changeMe!",
                    ForceChangePasswordNextLogin = true
                };
                user.PasswordProfile = PasswordProfile;
                user.PasswordPolicies = "DisablePasswordExpiration, DisableStrongPassword";
                try
                {
                    await user.UpdateAsync();
                    Console.WriteLine("\nUser {0} password and policy was reset", user.DisplayName);
                }
                catch (Exception e)
                {
                    Program.WriteError("\nError Updating the user {0}", Program.ExtractErrorMessage(e));
                }
            }

        }

        private static async Task AssignManager(IUser user, IUser manager)
        {
            // *************************************************************
            // These operations require more privileged permissions like Directory.ReadWrite.All or Directory.AccessAsUser.All
            // Update signed in user's manager, update group membership
            // **************************************************************


            // Assign the newly created user a new manager (the signed in user).
            if (user.ObjectId != null)
            {
                Console.WriteLine("\n Assigning {0} as {1}'s Manager.", manager.DisplayName,
                    user.DisplayName);
                user.Manager = manager as DirectoryObject;
                try
                {
                    await user.UpdateAsync();
                    Console.Write("Manager assignment successful.");
                }
                catch (Exception e)
                {
                    Program.WriteError("\nError assigning manager to user. {0}", Program.ExtractErrorMessage(e));
                }
            }

        }

        private static async Task AssignLicenses(IActiveDirectoryClient client, IUser user)
        {

            //*********************************************************************************************
            // User License Assignment - assign EnterprisePack license to new user, and disable SharePoint service
            //   first get a list of Tenant's subscriptions and find the "Enterprisepack" one
            //   Enterprise Pack includes service Plans for ExchangeOnline, SharePointOnline and LyncOnline
            //   validate that Subscription is Enabled and there are enough units left to assign to users
            //*********************************************************************************************
            IPagedCollection<ISubscribedSku> skus = null;
            try
            {
                skus = await client.SubscribedSkus.ExecuteAsync();
            }
            catch (Exception e)
            {
                Program.WriteError("\nError getting Applications {0}", Program.ExtractErrorMessage(e));
            }
            while (skus != null)
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
                            AssignedLicense addLicense = new AssignedLicense {SkuId = sku.SkuId.Value};

                            // find plan id of SharePoint Service Plan
                            foreach (ServicePlanInfo servicePlan in sku.ServicePlans)
                            {
                                if (servicePlan.ServicePlanName.Contains("SHAREPOINT"))
                                {
                                    addLicense.DisabledPlans.Add(servicePlan.ServicePlanId.Value);
                                    break;
                                }
                            }

                            IList<AssignedLicense> licensesToAdd = new[] {addLicense};
                            IList<Guid> licensesToRemove = new Guid[] {};

                            // attempt to assign the license object to the new user 
                            try
                            {
                                if (user.ObjectId != null)
                                {
                                    await user.AssignLicenseAsync(licensesToAdd, licensesToRemove);
                                    Console.WriteLine("\n User {0} was assigned license {1}",
                                        user.DisplayName,
                                        addLicense.SkuId);
                                }
                            }
                            catch (Exception e)
                            {
                                Program.WriteError("\nError Assigning License {0}", Program.ExtractErrorMessage(e));
                            }
                        }
                    }
                }
                skus = await skus.GetNextPageAsync();
            }

        }

        #endregion

        #region group operations

        private static async Task AddUserToGroup(Group group, IUser user)
        {
            // add new user to picked group
            try
            {
                group.Members.Add(user as DirectoryObject);
                await group.UpdateAsync();
            }
            catch (Exception e)
            {
                Program.WriteError("\nError assigning member to group. {0}", e.Message,
                    Program.ExtractErrorMessage(e));
            }
        }

        private static async Task PrintGroupMembers(ActiveDirectoryClient client)
        {
            // Search for a group and assign the newUser to the found group
            Console.WriteLine("\nSearch for a group, by name, to add the current user to:");
            string groupName = Console.ReadLine();
            //*********************************************************************
            // Search for a group using a startsWith filter (displayName property)
            //*********************************************************************
            List<IGroup> foundGroups = null;
            IGroup retrievedGroup = null;
            try
            {
                IPagedCollection<IGroup> groupsCollection = await client.Groups
                    .Where(g => g.DisplayName.StartsWith(groupName))
                    .ExecuteAsync();
                foundGroups = groupsCollection.CurrentPage.ToList();
            }
            catch (Exception e)
            {
                Program.WriteError("\nError getting Group {0}", e.Message, Program.ExtractErrorMessage(e));
            }
            if (foundGroups != null && foundGroups.Count > 0)
            {
                if (foundGroups.Count == 1)
                {
                    retrievedGroup = foundGroups[0];
                }
                else
                {
                    for (int i = 0; i < foundGroups.Count; i++)
                    {
                        Console.WriteLine("\n{0}. {1}", i + 1, foundGroups[i].DisplayName);
                    }

                    string keyString;
                    int key;
                    do
                    {
                        Console.WriteLine("Pick the group you want to add the new user to by entering a number:");
                        keyString = Console.ReadLine();
                    } while (!(int.TryParse(keyString, out key) && key > 0 && key <= foundGroups.Count));
                    retrievedGroup = foundGroups[key - 1];
                }
            }
            if (retrievedGroup != null && retrievedGroup.ObjectId != null)
            {
                Console.WriteLine("\n Enumerating group members for: " + retrievedGroup.DisplayName + "\n " +
                                  retrievedGroup.Description);

                //*********************************************************************
                // get the groups' membership - 
                // Note this method retrieves ALL links in one request - please use this method with care - this
                // may return a very large number of objects
                //*********************************************************************
                IGroupFetcher retrievedGroupFetcher = (IGroupFetcher) retrievedGroup;
                try
                {
                    IPagedCollection<IDirectoryObject> members = await retrievedGroupFetcher.Members.ExecuteAsync();
                    Console.WriteLine(" Members:");
                    do
                    {
                        List<IDirectoryObject> directoryObjects = members.CurrentPage.ToList();
                        foreach (IDirectoryObject member in directoryObjects)
                        {
                            if (member is User)
                            {
                                Console.WriteLine("User DisplayName: {0}  UPN: {1}",
                                    (member as User).DisplayName,
                                    (member as User).UserPrincipalName);
                            }
                            if (member is Group)
                            {
                                Console.WriteLine("Group DisplayName: {0}", (member as Group).DisplayName);
                            }
                            if (member is Contact)
                            {
                                Console.WriteLine("Contact DisplayName: {0}", (member as Contact).DisplayName);
                            }
                        }
                        members = await members.GetNextPageAsync();
                    } while (members != null);
                }
                catch (Exception e)
                {
                    Program.WriteError("\nError getting groups' membership. {0}", e.Message,
                        Program.ExtractErrorMessage(e));
                }
            }

        }

        private static async Task RemoveUserFromGroup(Group group, IUser user)
        {

            //*********************************************************************************************
            // Delete user from the earlier selected Group 
            //*********************************************************************************************
            if (group.ObjectId != null)
            {
                try
                {
                    group.Members.Remove(user as DirectoryObject);
                    await group.UpdateAsync();
                }
                catch (Exception e)
                {
                    Program.WriteError("\nError removing user from group {0}", Program.ExtractErrorMessage(e));
                }
            }

        }

        private static async Task<Group> CreateNewGroup(IActiveDirectoryClient client)
        {

            //*********************************************************************************************
            // Create a new Group
            //*********************************************************************************************
            Group newGroup = new Group
            {
                DisplayName = "newGroup" + Helper.GetRandomString(8),
                Description = "Best Group ever",
                MailNickname = "group" + Helper.GetRandomString(4),
                MailEnabled = false,
                SecurityEnabled = true
            };
            try
            {
                await client.Groups.AddGroupAsync(newGroup);
                Console.WriteLine("\nNew Group {0} was created", newGroup.DisplayName);
            }
            catch (Exception e)
            {
                Program.WriteError("\nError creating new Group {0}", Program.ExtractErrorMessage(e));
            }
            return newGroup;

        }

        #endregion

        #region Applications and Service Principals

        private static async Task<string> GetServicePrincipalObjectId(IActiveDirectoryClient client,
            string applicationId)
        {
            IPagedCollection<IServicePrincipal> servicePrincipals = null;
            try
            {
                servicePrincipals = await client.ServicePrincipals.ExecuteAsync();
            }
            catch (Exception e)
            {
                Program.WriteError("\nError getting Service Principal {0}", e.Message, Program.ExtractErrorMessage(e));
            }

            while (servicePrincipals != null)
            {
                List<IServicePrincipal> servicePrincipalsList = servicePrincipals.CurrentPage.ToList();
                IServicePrincipal servicePrincipal =
                    servicePrincipalsList.FirstOrDefault(x => x.AppId.Equals(applicationId));

                if (servicePrincipal != null)
                {
                    return servicePrincipal.ObjectId;
                }

                servicePrincipals = await servicePrincipals.GetNextPageAsync();
            }
            return string.Empty;
        }

        private static async Task PrintApplications(IActiveDirectoryClient client)
        {

            //*********************************************************************
            // get the Application objects
            //*********************************************************************
            IPagedCollection<IApplication> applications = null;
            try
            {
                applications = await client.Applications.Take(50).ExecuteAsync();
            }
            catch (Exception e)
            {
                Program.WriteError("\nError getting Applications {0}", Program.ExtractErrorMessage(e));
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
                    applications = await applications.GetNextPageAsync();
                } while (applications != null);
            }

        }

        private static async Task PrintServicePrincipals(IActiveDirectoryClient client)
        {

            //*********************************************************************
            // get the Service Principals
            //*********************************************************************
            IPagedCollection<IServicePrincipal> servicePrincipals = null;
            try
            {
                servicePrincipals = await client.ServicePrincipals.ExecuteAsync();
            }
            catch (Exception e)
            {
                Program.WriteError("\nError getting Service Principal {0}", e.Message, Program.ExtractErrorMessage(e));
            }
            while (servicePrincipals != null)
            {
                List<IServicePrincipal> servicePrincipalsList = servicePrincipals.CurrentPage.ToList();
                foreach (IServicePrincipal servicePrincipal in servicePrincipalsList)
                {
                    Console.WriteLine("Service Principal AppId: {0}  Name: {1}", servicePrincipal.AppId,
                        servicePrincipal.DisplayName);
                }
                servicePrincipals = await servicePrincipals.GetNextPageAsync();
            }

        }

        private static async Task<Application> CreateNewApplication(IActiveDirectoryClient client, IUser user)
        {

            //*********************************************************************************************
            // Create a new Application object, with an App Role definition
            //*********************************************************************************************
            Application newApp = new Application {DisplayName = "Test-Demo App " + Helper.GetRandomString(4)};
            newApp.IdentifierUris.Add("https://localhost/demo/" + Guid.NewGuid());
            newApp.ReplyUrls.Add("https://localhost/demo");
            AppRole appRole = new AppRole()
            {
                Id = Guid.NewGuid(),
                IsEnabled = true,
                DisplayName = "Something",
                Description = "Anything",
                Value = "policy.write"
            };
            appRole.AllowedMemberTypes.Add("User");
            newApp.AppRoles.Add(appRole);

            // Add a password key
            PasswordCredential password = new PasswordCredential
            {
                StartDate = DateTime.UtcNow,
                EndDate = DateTime.UtcNow.AddYears(1),
                Value = "password",
                KeyId = Guid.NewGuid()
            };
            newApp.PasswordCredentials.Add(password);

            try
            {
                await client.Applications.AddApplicationAsync(newApp);
                Console.WriteLine("New Application created: " + newApp.DisplayName);
            }
            catch (Exception e)
            {
                Program.WriteError("\nError ceating Application: {0}", Program.ExtractErrorMessage(e));
            }

            // add an owner for the newly created application
            newApp.Owners.Add(user as DirectoryObject);
            try
            {
                await newApp.UpdateAsync();
                Console.WriteLine("Added owner: " + newApp.DisplayName, user.DisplayName);
            }
            catch (Exception e)
            {
                Program.WriteError("\nError adding Application owner: {0}", Program.ExtractErrorMessage(e));
            }

            // check the ownership for the newly created application
            try
            {
                IApplication appCheck = await client.Applications.GetByObjectId(newApp.ObjectId).ExecuteAsync();
                IApplicationFetcher appCheckFetcher = appCheck as IApplicationFetcher;

                IPagedCollection<IDirectoryObject> appOwners = await appCheckFetcher.Owners.ExecuteAsync();

                do
                {
                    List<IDirectoryObject> directoryObjects = appOwners.CurrentPage.ToList();
                    foreach (IDirectoryObject directoryObject in directoryObjects)
                    {
                        if (directoryObject is User)
                        {
                            User appOwner = directoryObject as User;
                            Console.WriteLine("Application {0} has {1} as owner", appCheck.DisplayName,
                                appOwner.DisplayName);
                        }
                    }
                    appOwners = await appOwners.GetNextPageAsync();
                } while (appOwners != null);
            }
            catch (Exception e)
            {
                Program.WriteError("\nError checking Application owner: {0}", Program.ExtractErrorMessage(e));
            }
            return newApp;

        }

        private static async Task<ServicePrincipal> CreateServicePrincipal(IActiveDirectoryClient client,
            IApplication application)
        {

            //*********************************************************************************************
            // create a new Service principal, from the application object that was just created
            //*********************************************************************************************
            ServicePrincipal newServicePrincipal = new ServicePrincipal();
            if (application.AppId != null)
            {
                newServicePrincipal.DisplayName = application.DisplayName;
                newServicePrincipal.AccountEnabled = true;
                newServicePrincipal.AppId = application.AppId;
                try
                {
                    await client.ServicePrincipals.AddServicePrincipalAsync(newServicePrincipal);
                    Console.WriteLine("New Service Principal created: " + newServicePrincipal.DisplayName);
                }
                catch (Exception e)
                {
                    Program.WriteError("\nError Creating Service Principal: {0}", Program.ExtractErrorMessage(e));
                }
            }
            return newServicePrincipal;
        }

        #endregion

        #region Extension Properties

        private static void PrintExtensionProperty(User user, ExtensionProperty extension)
        {

            try
            {
                if (extension != null && user != null && user.ObjectId != null)
                {
                    Console.WriteLine("\n Retrieved User {0}'s extended property value is: {1}.",
                        user.DisplayName,
                        extension);
                }
            }
            catch (Exception e)
            {
                Program.WriteError("\nError Updating the user object {0}", Program.ExtractErrorMessage(e));
            }

        }

        private static async Task<ExtensionProperty> CreateSchemaExtensions(
            IActiveDirectoryClient client, Application application)
        {

            // **************************************************************************************************
            // Create a new extension property - to extend the user entity
            // This is accomplished by declaring the extension property through an application object
            // **************************************************************************************************
            if (application.ObjectId != null)
            {
                ExtensionProperty linkedInUserId = new ExtensionProperty
                {
                    Name = "linkedInUserId",
                    DataType = "String",
                    TargetObjects = {"User"}
                };
                try
                {
                    // firstly, let's write out all the existing cloud extension properties in the tenant
                    IEnumerable<IExtensionProperty> allExts = await client.GetAvailableExtensionPropertiesAsync(false);
                    foreach (ExtensionProperty ext in allExts)
                    {
                        Console.WriteLine("\nExtension: {0}", ext.Name);
                    }
                    application.ExtensionProperties.Add(linkedInUserId);
                    await application.UpdateAsync();
                    Console.WriteLine("\nUser object extended successfully with extension: {0}.", linkedInUserId.Name);
                }
                catch (Exception e)
                {
                    Program.WriteError("\nError extending the user object {0}", Program.ExtractErrorMessage(e));
                }
                return linkedInUserId;

            }
            return null;
        }

        private static async Task SetExtensionProperty(ExtensionProperty extension, User user)
        {

            // **************************************************************************************************
            // Update an extension property that exists on the user entity
            // **************************************************************************************************

            // create the extension attribute name, for the extension that we just created
            try
            {
                if (extension != null && user != null && user.ObjectId != null)
                {
                    user.SetExtendedProperty(extension.Name, "user@linkedin.com");
                    await user.UpdateAsync();
                    Console.WriteLine("\nUser {0}'s extended property set successully.", user.DisplayName);
                }
            }
            catch (Exception e)
            {
                Program.WriteError("\nError Updating the user object {0}", Program.ExtractErrorMessage(e));
            }

        }

        #endregion

        #region Permissions and Roles

        private static async Task AssignAppRole(IActiveDirectoryClient client,
            Application application,
            ServicePrincipal servicePrincipal)
        {

            try
            {
                User user =
                    (User) (await client.Users.ExecuteAsync()).CurrentPage.ToList().FirstOrDefault();
                if (application.ObjectId != null && user != null && servicePrincipal.ObjectId != null)
                {
                    // create the app role assignment
                    AppRoleAssignment appRoleAssignment = new AppRoleAssignment();
                    appRoleAssignment.Id = application.AppRoles.FirstOrDefault().Id;
                    appRoleAssignment.ResourceId = Guid.Parse(servicePrincipal.ObjectId);
                    appRoleAssignment.PrincipalType = "User";
                    appRoleAssignment.PrincipalId = Guid.Parse(user.ObjectId);
                    user.AppRoleAssignments.Add(appRoleAssignment);

                    // assign the app role
                    await user.UpdateAsync();
                    Console.WriteLine("User {0} is successfully assigned an app (role).", user.DisplayName);

                    // remove the app role
                    user.AppRoleAssignments.Remove(appRoleAssignment);
                    await user.UpdateAsync();
                    Console.WriteLine("User {0} is successfully removed from app (role).", user.DisplayName);

                }
            }

            catch (Exception e)
            {
                Program.WriteError("\nError Assigning Direct Permission: {0}", Program.ExtractErrorMessage(e));
            }

        }

        private static async Task PrintDevicesAndOwners(IActiveDirectoryClient client)
        {

            //*********************************************************************************************
            // Get a list of Mobile Devices from tenant
            //*********************************************************************************************
            Console.WriteLine("\nGetting Devices");
            IPagedCollection<IDevice> devices = null;
            try
            {
                devices = await client.Devices.ExecuteAsync();
            }
            catch (Exception e)
            {
                Program.WriteError("/nError getting devices {0}", Program.ExtractErrorMessage(e));
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
                                    registeredOwners = await registeredOwners.GetNextPageAsync();
                                } while (registeredOwners != null);
                            }
                        }
                    }
                    devices = await devices.GetNextPageAsync();
                } while (devices != null);
            }

        }

        private static async Task PrintAllRoles(IActiveDirectoryClient client, string searchString)
        {

            //*********************************************************************
            // Get All Roles
            //*********************************************************************
            List<IDirectoryRole> foundRoles = null;
            try
            {
                IPagedCollection<IDirectoryRole> rolesCollection = await client.DirectoryRoles.ExecuteAsync();
                foundRoles = rolesCollection.CurrentPage.ToList();
            }
            catch (Exception e)
            {
                Program.WriteError("\nError getting Roles {0}", Program.ExtractErrorMessage(e));
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

        }

        private static async Task PrintAllPermissions(IActiveDirectoryClient client)
        {

            //*********************************************************************************************
            // get all Permission Objects
            //*********************************************************************************************
            Console.WriteLine("\n Getting Permissions");
            IPagedCollection<IOAuth2PermissionGrant> permissions = null;
            try
            {
                permissions = await client.Oauth2PermissionGrants.ExecuteAsync();
            }
            catch (Exception e)
            {
                Program.WriteError("\nError Getting Permissions: {0}", Program.ExtractErrorMessage(e));
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
                    permissions = await permissions.GetNextPageAsync();
                } while (permissions != null);
            }

        }

        private static async Task<OAuth2PermissionGrant> CreateOAuth2Permission(
            IActiveDirectoryClient client,
            ServicePrincipal servicePrincipal)
        {

            //*********************************************************************************************
            // Create new  oauth2 permission object
            //*********************************************************************************************
            if (servicePrincipal.ObjectId != null)
            {
                OAuth2PermissionGrant permissionObject = new OAuth2PermissionGrant
                {
                    ConsentType = "AllPrincipals",
                    Scope = "user_impersonation",
                    StartTime = DateTime.Now,
                    ExpiryTime = (DateTime.Now).AddMonths(12),
                    ResourceId = await GetServicePrincipalObjectId(client, GlobalConstants.GraphServiceObjectId),
                    ClientId = servicePrincipal.ObjectId
                };
                try
                {
                    await client.Oauth2PermissionGrants.AddOAuth2PermissionGrantAsync(permissionObject);
                    Console.WriteLine("New Permission object created: " + permissionObject.ObjectId);
                }
                catch (Exception e)
                {
                    Program.WriteError("\nError with Permission Creation: {0}", Program.ExtractErrorMessage(e));
                }

                return permissionObject;
            }


            return null;
        }

        #endregion

        #region Domain Operations

        private static async Task PrintAllDomains(IActiveDirectoryClient client)
        {
            //*********************************************************************************************
            // get all Domains
            //*********************************************************************************************
            Console.WriteLine("\n Getting Domains");
            IPagedCollection<IDomain> domains = null;
            try
            {
                domains = await client.Domains.ExecuteAsync();
            }
            catch (Exception e)
            {
                Program.WriteError("\nError Getting Domains: {0}", Program.ExtractErrorMessage(e));
            }
            while (domains != null)
            {
                List<IDomain> domainList = domains.CurrentPage.ToList();
                foreach (IDomain domain in domainList)
                {
                    Console.WriteLine("Domain: {0}  Verified: {1}", domain.Name, domain.IsVerified);
                }
                domains = await domains.GetNextPageAsync();
            }
        }

        private static async Task<IDomain> CreateNewDomain(IActiveDirectoryClient client)
        {
            IDomain newDomain = new Domain {Name = Helper.GetRandomString() + ".com"};
            try
            {
                await client.Domains.AddDomainAsync(newDomain);
                Console.WriteLine("\nNew Domain {0} was created", newDomain.Name);
            }
            catch (Exception e)
            {
                Program.WriteError("\nError creating new Domain {0}", Program.ExtractErrorMessage(e));
            }
            return newDomain;
        }

        private static void PrintDomainVerificationDetails(IDomainFetcher domainFetcher)
        {
            // get verification details - this onfo is used to update your registrar/DNS host, so that verification can be performed
            try
            {
                IPagedCollection<IDomainDnsRecord> verificationRecords =
                    domainFetcher.VerificationDnsRecords.ExecuteAsync().Result;
                List<IDomainDnsRecord> records = verificationRecords.CurrentPage.ToList();
                // Normally would page through these, but it shouldn't go over a page, so no need
                foreach (IDomainDnsRecord record in records)
                {
                    if (record is DomainDnsTxtRecord)
                    {
                        DomainDnsTxtRecord txtRecord = record as DomainDnsTxtRecord;
                        Console.WriteLine("TXT Record:\n  Label: {0}\n  VerifyText: {1}",
                            txtRecord.Label, txtRecord.Text);
                    }
                    if (record is DomainDnsMxRecord)
                    {
                        DomainDnsMxRecord txtRecord = record as DomainDnsMxRecord;
                        Console.WriteLine("MX Record:\n  Label: {0}\n  VerifyText: {1}\n  Preference: {2}",
                            txtRecord.Label, txtRecord.MailExchange, txtRecord.Preference);
                    }
                }
            }
            catch (Exception e)
            {
                Program.WriteError("\nError getting verification records {0}", Program.ExtractErrorMessage(e));
            }

        }

        private static void VerifyDomain(IDomain domain)
        {
            // now attempt to verify the domain
            // this should be run in a loop with a backoof delay, to keep calling verify until the domain is verified.
            int count = 0;
            try
            {
                while (!domain.IsVerified && count < 5)
                {
                    count++;
                    domain.VerifyAsync().Wait();
                }

            }
            catch (Exception e)
            {
                Program.WriteError("\nError verifying domain {0}", Program.ExtractErrorMessage(e));
                Thread.Sleep(1000*count);
            }
        }

        #endregion

        #region CleanUp

        private static async Task DeleteUser(IUser user)
        {
            #region Delete user

            //*********************************************************************************************
            // Delete the user that we just created earlier
            //*********************************************************************************************
            if (user != null)
            {
                try
                {
                    await user.DeleteAsync();
                    Console.WriteLine("\nUser {0} was deleted", user.DisplayName);
                }
                catch (Exception e)
                {
                    Program.WriteError("\nError Deleting User {0}", Program.ExtractErrorMessage(e));
                }
            }

            #endregion
        }

        private static async Task DeleteGroup(Group group)
        {
            #region Delete Group

            //*********************************************************************************************
            // Delete the Group that we just created
            //*********************************************************************************************
            if (group != null)
            {
                try
                {
                    await group.DeleteAsync();
                    Console.WriteLine("\nGroup {0} was deleted", group.DisplayName);
                }
                catch (Exception e)
                {
                    Program.WriteError("\nError Deleting Group {0}", Program.ExtractErrorMessage(e));
                }
            }

            #endregion
        }

        private static async Task DeleteApplication(Application application, ExtensionProperty newExtension = null)
        {
            #region Delete Application

            //*********************************************************************************************
            // Delete Application Objects
            //*********************************************************************************************
            if (application != null)
            {
                try
                {
                    if (newExtension != null)
                    {
                        try
                        {
                            application.ExtensionProperties.Remove(newExtension);
                            Console.WriteLine("\nDeleted extension property: " + newExtension.Name);
                        }
                        catch (Exception e)
                        {
                            Program.WriteError("\nError deleting extension property: {0}",
                                Program.ExtractErrorMessage(e));
                        }
                    }
                    await application.DeleteAsync();
                    Console.WriteLine("\nDeleted Application object: " + application.ObjectId);
                }
                catch (Exception e)
                {
                    Program.WriteError("\nError deleting Application: {0}", Program.ExtractErrorMessage(e));
                }
            }

            #endregion
        }

        private static async Task DeleteServicePrincipalAndPermission(
            ServicePrincipal servicePrincipal, 
            OAuth2PermissionGrant permissionGrant = null)
        {
            if (servicePrincipal != null)
            {
                try
                {
                    // remove the oauth2 permission scope grant
                    if (permissionGrant != null)
                    {
                        try
                        {
                            servicePrincipal.Oauth2PermissionGrants.Remove(permissionGrant);
                            await servicePrincipal.UpdateAsync();
                            Console.WriteLine("Removed Permission object: " + permissionGrant.ObjectId);
                        }

                        catch (Exception e)
                        {
                            Program.WriteError("\nError with Permission Deletion: {0}", Program.ExtractErrorMessage(e));
                        }
                    }
                    await servicePrincipal.DeleteAsync();
                    Console.WriteLine("Deleted service principal object: " + servicePrincipal.ObjectId);                
                }
                catch (Exception e)
                {
                    Program.WriteError("\nError with Service Principal deletion: {0}", Program.ExtractErrorMessage(e));
                }
            }
        }

        private static async Task DeleteDomain(IDomain domain)
        {
            #region Delete Domain

            //*********************************************************************************************
            // Delete Domain we created
            //*********************************************************************************************
            if (domain != null)
            {
                try
                {
                    await domain.DeleteAsync();
                    Console.WriteLine("\nDeleted Domain: " + domain.Name);
                }
                catch (Exception e)
                {
                    Program.WriteError("\nError deleting Domain: {0}", Program.ExtractErrorMessage(e));
                }
            }

            #endregion
        }

        #endregion

        #region batch operations
        private static async Task BatchOperations(ActiveDirectoryClient client)
        {
            //*********************************************************************************************
            // Show Batching with 3 operators.  Note: up to 5 operations can be in a batch
            //*********************************************************************************************
            IReadOnlyQueryableSet<User> userQuery = client.DirectoryObjects.OfType<User>();
            IReadOnlyQueryableSet<Group> groupsQuery = client.DirectoryObjects.OfType<Group>();
            IReadOnlyQueryableSet<DirectoryRole> rolesQuery =
                client.DirectoryObjects.OfType<DirectoryRole>();
            try
            {
                IBatchElementResult[] batchResult = await
                    client.Context.ExecuteBatchAsync(userQuery, groupsQuery, rolesQuery);
                int responseCount = 1;
                foreach (IBatchElementResult result in batchResult)
                {
                    if (result.FailureResult != null)
                    {
                        Console.WriteLine("Batch Item Result {0} failed. Exception: {1} ",
                            responseCount, result.FailureResult.InnerException);
                    }
                    else if (result.SuccessResult != null)
                    {
                        Console.WriteLine("Batch Item Result {0} succeeded",
                            responseCount);
                    }
                    responseCount++;
                }
            }
            catch (Exception e)
            {
                Program.WriteError("\nError with batch execution. : {0}", Program.ExtractErrorMessage(e));
            }
        }
        #endregion
    }
}