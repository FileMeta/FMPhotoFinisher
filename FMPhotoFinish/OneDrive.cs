using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.IO;
using System.Xml.XPath;

namespace FMPhotoFinish
{

    // Useful documentation links:
    // https://docs.microsoft.com/en-us/onedrive/developer/rest-api/api/drive_get_specialfolder

    class OneDrive
    {
        string m_accessToken;

        public OneDrive(string accessToken)
        {
            m_accessToken = accessToken;
        }

        public void Test()
        {
            DumpJson(FetchJson("https://graph.microsoft.com/v1.0/me/drive/special/cameraroll/children"));
        }

        string Fetch(string url)
        {
            using (var response = FetchResponse(url))
            {
                using (var reader = new StreamReader(response.GetResponseStream()))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        XPathDocument FetchJson(string url)
        {
            using (var response = FetchResponse(url))
            {
                using (var reader = System.Runtime.Serialization.Json.JsonReaderWriterFactory.CreateJsonReader(response.GetResponseStream(), new System.Xml.XmlDictionaryReaderQuotas()))
                {
                    return new XPathDocument(reader);
                }
            }
        }

        HttpWebResponse FetchResponse(string url)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.Headers.Add(HttpRequestHeader.Authorization, "Bearer " + m_accessToken);
            try
            {
                return (HttpWebResponse)request.GetResponse();
            }
            catch (WebException ex)
            {
                string errTxt = string.Empty;
                var response = ex.Response as HttpWebResponse;
                if (response != null)
                {
                    using (var reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
                    {
                        errTxt = reader.ReadToEnd();
                    }
                }

                throw new Exception("Http Error: " + errTxt);
            }

        }

        void DumpJson(XPathDocument json)
        {
            using (var writer = System.Runtime.Serialization.Json.JsonReaderWriterFactory
                .CreateJsonWriter(Console.OpenStandardOutput(), new UTF8Encoding(false), true, true))
            {
                json.CreateNavigator().WriteSubtree(writer);
            }
        }
    }
}
