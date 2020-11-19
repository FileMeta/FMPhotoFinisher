using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using Windows.Security.Credentials;

/* To add the reference for PasswordVault in Visual Studio:
 * 1. Add a reference
 * 2. Browse to C:\Program Files (x86)\Windows Kits\8.1\References\CommonConfiguration\Neutral\Annotated\Windows.winmd"
 * (Note that the .winmd extension is not selected by default)
 */


namespace FileMeta
{
    // Access Microsoft OneDrive
    class NamedSource
    {
        const string c_pvResource = "FMPhotoFinish";
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
                // Store the refresh token in the password vault
                var vault = new PasswordVault();
                vault.Add(new PasswordCredential(c_pvResource, sourceName, c_oneDrivePrefix + oa.Refresh_Token));
                Console.WriteLine($"Successfully created OneDrive named source: {sourceName}");
            }
            else
            {
                Console.WriteLine("OneDrive authorization failure.");
            }
        }

        public static void TestAccess(string sourceName)
        {
            var vault = new PasswordVault();
            var cred = vault.Retrieve("FMPhotoFinish", sourceName);

            var password = cred.Password;
            if (!password.StartsWith(c_oneDrivePrefix, StringComparison.OrdinalIgnoreCase))
            {
                throw new Exception("Named source is not 'OneDrive:'");
            }

            var oa = new OAuth(c_microsoftAuthorizaitonEndpoint, c_microsoftTokenExchangeEndpoint, c_clientId);
            if (!oa.Refresh(password.Substring(c_oneDrivePrefix.Length)))
            {
                throw new Exception($"Failed to refresh OneDrive access: {oa.Error}");
            }

            Console.WriteLine(oa.Access_Token);
        }

    }
}
