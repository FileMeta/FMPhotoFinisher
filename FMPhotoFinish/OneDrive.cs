using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;


namespace FileMeta
{
    // Access Microsoft OneDrive
    class OneDrive
    {
        const string c_microsoftAuthorizaitonEndpoint = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";
        const string c_microsoftTokenExchangeEndpoint = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
        const string c_clientId = "78347c9d-b79a-4276-a772-f6563b012d70";
        const string c_scopes = "Files.Read offline_access openid User.Read";

        public static void LoginAndAuthorize(string sourceName)
        {
            var oa = new OAuth(c_microsoftAuthorizaitonEndpoint, c_microsoftTokenExchangeEndpoint,
                c_clientId);
            if (oa.Authorize(c_scopes))
            {
                Debug.WriteLine($"access_token={oa.Access_Token} refresh_token={oa.Refresh_Token}");
            }
        }

    }
}
