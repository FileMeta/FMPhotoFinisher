using System;
using FileMeta;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.XPath;
using System.Net;
using System.IO;


namespace FMPhotoFinish
{
    class OneDriveSource : IMediaSource
    {
        const string c_targetPrefix = "FMPhotoFinish:";
        const string c_oneDriveBookmarkPrefix = "OneDrive/";
        const string c_oneDriveCameraRollUrl = @"https://graph.microsoft.com/v1.0/me/drive/special/cameraroll/children";

        string m_sourceName;
        string m_refreshToken;
        string m_accessToken;
        string m_bookmarkPath;

        public OneDriveSource(string sourceName, string refreshToken)
        {
            m_sourceName = sourceName;
            m_refreshToken = refreshToken;
            m_bookmarkPath = c_oneDriveBookmarkPrefix + sourceName;
        }

        public void RetrieveMediaFiles(SourceConfiguration sourceConfig, IMediaQueue mediaQueue)
        {
            mediaQueue.ReportProgress($"Connecting to OneDrive: {m_sourceName}");
            m_accessToken = NamedSource.GetOnedriveAccessToken(m_refreshToken);
            var queue = SelectFiles(sourceConfig, mediaQueue);
        }

        List<string> SelectFiles(SourceConfiguration sourceConfig, IMediaQueue mediaQueue)
        {
            mediaQueue.ReportProgress("Selecting files");

            // Determine the "after" threshold from SelectAfter and SelectIncremental
            var after = sourceConfig.GetBookmarkOrAfter(m_bookmarkPath);
            if (after.HasValue)
            {
                mediaQueue.ReportProgress($"  Filter: Created after {after.Value}");
            }

            // Retrieve the filenames
            int count = 0;
            string nextUrl = c_oneDriveCameraRollUrl;
            for (; ; )
            {
                Console.WriteLine(nextUrl);
                var fileList = FetchJson(nextUrl);
                var root = fileList.CreateNavigator();
                nextUrl = root.XPVal("/root/a:item[@item='@odata.nextLink']");

                foreach (XPathNavigator node in fileList.CreateNavigator().Select("/root/value/item"))
                {
                    Console.WriteLine(node.XPVal("name"));
                    ++count;
                }

                if (string.IsNullOrEmpty(nextUrl)) break;
            }
            Console.WriteLine($"{count} files");
            return null;
        }


        #region REST Client

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

        void DumpJson(XPathDocument json, string filename)
        {
            using (var writer = System.Runtime.Serialization.Json.JsonReaderWriterFactory
                .CreateJsonWriter(new FileStream(filename, FileMode.Create, FileAccess.Write, FileShare.None),
                    new UTF8Encoding(false), true, true))
            {
                json.CreateNavigator().WriteSubtree(writer);
            }
        }

        void DumpXml(XPathDocument json, string filename)
        {
            var settings = new System.Xml.XmlWriterSettings();
            settings.Indent = true;
            using (var writer = System.Xml.XmlWriter.Create(filename, settings))
            {
                json.CreateNavigator().WriteSubtree(writer);
            }
        }

        #endregion
    }

    static class XmlHelp
    {
        public static readonly XmlNamespaceManager Xns;

        static XmlHelp()
        {
            Xns = new XmlNamespaceManager(new NameTable());
            Xns.AddNamespace("a", "item");
        }

        public static string XPVal(this XPathNavigator nav, string xpath)
        {
            return nav.SelectSingleNode(xpath, Xns)?.Value;
        }
    }
}
