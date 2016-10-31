namespace GraphConsoleAppV3
{
    internal class AppModeConstants
    {
        public const string ClientId = "118473c2-7619-46e3-a8e4-6da8d5f56e12";
        public const string ClientSecret = "tdzicjqRJ2R1Hnvn3AL1aDU+yQjX5oIHz0th0qBeOOI=";
        public const string TenantName = "GraphDir1.onMicrosoft.com";
        public const string TenantId = "4fd2b2f2-ea27-4fe5-a8f3-7b1a7c975f34";
        public const string AuthString = GlobalConstants.AuthString + TenantName;
    }

    internal class UserModeConstants
    {
        public const string TenantId = AppModeConstants.TenantId;
        public const string ClientId = "66133929-66a4-4edc-aaee-13b04b03207d";
        public const string AuthString = GlobalConstants.AuthString + "common/";
    }

    internal class GlobalConstants
    {
        public const string AuthString = "https://login.microsoftonline.com/";        
        public const string ResourceUrl = "https://graph.windows.net";
        public const string GraphServiceObjectId = "00000002-0000-0000-c000-000000000000";
    }
}
