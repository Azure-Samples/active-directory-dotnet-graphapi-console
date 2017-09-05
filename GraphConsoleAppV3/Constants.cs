namespace GraphConsoleAppV3
{
    internal class AppModeConstants
    {
        public const string ClientSecret = "";
        public const string TenantName = "";
        public const string AuthString = GlobalConstants.AuthString + TenantName;
    }

    internal class UserModeConstants
    {
        public const string AuthString = GlobalConstants.AuthString + "common/";
    }

    internal class GlobalConstants
    {
        public const string AuthString = "https://login.microsoftonline.com/";        
        public const string ResourceUrl = "https://graph.windows.net";
        public const string GraphServiceObjectId = "00000002-0000-0000-c000-000000000000";
        public const string TenantId = "";
        public const string ClientId = "";
    }
}
