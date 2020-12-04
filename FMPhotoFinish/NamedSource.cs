using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using FileMeta;

namespace FMPhotoFinish
{

    // OneDrive API Docs: https://docs.microsoft.com/en-us/onedrive/developer/rest-api/getting-started/

    // Access Microsoft OneDrive
    class NamedSource
    {
        const string c_targetPrefix = "FMPhotoFinish:";
        const string c_oneDrivePrefix = "OneDrive:";

        const string c_microsoftAuthorizaitonEndpoint = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";
        const string c_microsoftTokenExchangeEndpoint = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
        const string c_clientId = "78347c9d-b79a-4276-a772-f6563b012d70";
        const string c_scopes = "Files.Read offline_access openid User.Read"; // offline_access is required to get a refresh token

        public static void OneDriveLoginAndAuthorize(string sourceName)
        {
            var oa = new OAuth(c_microsoftAuthorizaitonEndpoint, c_microsoftTokenExchangeEndpoint, c_clientId);
            if (oa.Authorize(c_scopes))
            {
                // Store the refresh token in the Credential manager
                CredentialManager.Add(c_targetPrefix + sourceName, string.Empty, c_oneDrivePrefix + oa.Refresh_Token);
                Console.WriteLine($"Successfully created OneDrive named source: {sourceName}");
            }
            else
            {
                Console.WriteLine("OneDrive authorization failure.");
            }
        }

        public static void TestAccess(string sourceName)
        {
            string username;
            string credential;
            if (!CredentialManager.Retrieve(c_targetPrefix + sourceName, out username, out credential))
            {
                throw new Exception($"Named source not found: {sourceName}");
            }
            if (!credential.StartsWith(c_oneDrivePrefix, StringComparison.OrdinalIgnoreCase))
            {
                throw new Exception("Named source is not 'OneDrive:'");
            }

            var oa = new OAuth(c_microsoftAuthorizaitonEndpoint, c_microsoftTokenExchangeEndpoint, c_clientId);
            if (!oa.Refresh(credential.Substring(c_oneDrivePrefix.Length)))
            {
                throw new Exception($"Failed to refresh OneDrive access: {oa.Error}");
            }

            var od = new OneDrive(oa.Access_Token);
            od.Test();
        }



    }
}
