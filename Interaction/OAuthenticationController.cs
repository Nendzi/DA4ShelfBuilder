using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Forge;

namespace Interaction
{
    public class OAuthenticationController
    {
        private static dynamic InternalToken { get; set; }
        private static Scope[] scope =
        {
            Scope.BucketCreate,
            Scope.BucketRead,
            Scope.BucketDelete,
            Scope.DataRead,
            Scope.DataWrite,
            Scope.DataCreate,
            Scope.CodeAll
        };
        public static async Task<dynamic> GetInternalAsync()
        {
            if (InternalToken == null || InternalToken.ExpiresAt < DateTime.UtcNow)
            {
                InternalToken = await Get2LeggedTokenAsync(scope);
                InternalToken.ExpiresAt = DateTime.UtcNow.AddSeconds(InternalToken.expires_in);
            }
            return InternalToken;
        }
        private static async Task<dynamic> Get2LeggedTokenAsync(Scope[] scopes)
        {
            TwoLeggedApi oauth = new TwoLeggedApi();
            string grantType = "client_credentials";
            dynamic bearer = await oauth.AuthenticateAsync(
                GetAppSetting("FORGE_CLIENT_ID"),
                GetAppSetting("FORGE_CLIENT_SECRET"),
                grantType,
                scopes);
            return bearer;
        }
        public static string GetAppSetting(string settingKey)
        {
            var UsedID = Environment.GetEnvironmentVariable(settingKey);
            return UsedID;
        }
    }
}
