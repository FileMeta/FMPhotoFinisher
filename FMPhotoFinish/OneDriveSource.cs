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
using System.Globalization;


namespace FMPhotoFinish
{
    class OneDriveSource : IMediaSource
    {
        const string c_targetPrefix = "FMPhotoFinish:";
        const string c_oneDriveBookmarkPrefix = "OneDrive:";
        const string c_oneDriveCameraRollUrl = @"https://graph.microsoft.com/v1.0/me/drive/special/cameraroll/children";
        const string c_oneDriveItemPrefix = @"https://graph.microsoft.com/v1.0/me/drive/items/";    // Should be followed by the item ID
        const string c_oneDriveContentSuffix = @"/content";
        const int c_maxBatch = 150;

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
            // if a destination directory not specified, error
            if (string.IsNullOrEmpty(sourceConfig.DestinationDirectory))
                throw new InvalidOperationException("OneDrive Source requires destination directory -d");

            mediaQueue.ReportProgress($"Connecting to OneDrive: {m_sourceName}");
            m_accessToken = NamedSource.GetOnedriveAccessToken(m_refreshToken);
            var queue = SelectFiles(sourceConfig, mediaQueue);

            DownloadMediaFiles(queue, sourceConfig, mediaQueue);
        }

        List<OdFileInfo> SelectFiles(SourceConfiguration sourceConfig, IMediaQueue mediaQueue)
        {
            mediaQueue.ReportProgress("Selecting files");

            // Determine the "after" threshold from SelectAfter and SelectIncremental
            var after = sourceConfig.GetBookmarkOrAfter(m_bookmarkPath);
            if (after.HasValue)
            {
                mediaQueue.ReportProgress($"  Filter: Created after {after.Value}");
            }

            // Retrieve file info and enqueue
            var queue = new List<OdFileInfo>();
            int count = 0;
            string nextUrl = c_oneDriveCameraRollUrl;
            do
            {
                var fileList = FetchJson(nextUrl);
                var root = fileList.CreateNavigator();
                nextUrl = root.XPVal("/root/a:item[@item='@odata.nextLink']");

                foreach (XPathNavigator node in fileList.CreateNavigator().Select("/root/value/item"))
                {
                    ++count;
                    var odfi = new OdFileInfo(node);

                    if (!after.HasValue
                        || (odfi.BookmarkDate.HasValue && odfi.BookmarkDate > after.Value))
                    {
                        queue.Add(odfi);
                    }
                }

                mediaQueue.ReportStatus($"  Selected {queue.Count} Skipped {count - queue.Count}");
            }
            while (!string.IsNullOrEmpty(nextUrl));
            mediaQueue.ReportStatus(null);
            mediaQueue.ReportProgress($"Selected {queue.Count} Skipped {count - queue.Count}");

            if (queue.Count > c_maxBatch)
            {
                // Sort so that we'll keep the oldest ones.
                queue.Sort((a, b) => DateCompare(a.BookmarkDate, b.BookmarkDate));
                queue.RemoveRange(c_maxBatch, queue.Count - c_maxBatch);
                mediaQueue.RequestAnotherBatch = true;
                mediaQueue.ReportProgress($"Batch limited to {queue.Count}");
            }

            return queue;
        }

        static int DateCompare(DateTime? a, DateTime? b)
        {
            if (!a.HasValue && !b.HasValue) return 0;
            if (!a.HasValue) return 1;
            if (!b.HasValue) return -1;
            return a.Value.CompareTo(b.Value);
        }

        void DownloadMediaFiles(List<OdFileInfo> queue, SourceConfiguration sourceConfig, IMediaQueue mediaQueue)
        {
            mediaQueue.ReportProgress($"Downloading media files from OneDrive to working folder: {sourceConfig.DestinationDirectory}.");
            DateTime newestSelection = DateTime.MinValue;

            // Sum up the size of the files to be downloaded
            long selectedFilesSize = 0;
            foreach (var fi in queue)
            {
                selectedFilesSize += fi.Size;
            }

            uint startTicks = (uint)Environment.TickCount;
            long bytesDownloaded = 0;

            int n = 0;
            foreach (var fi in queue)
            {
                if (bytesDownloaded == 0)
                {
                    mediaQueue.ReportStatus($"Downloading file {n + 1} of {queue.Count}");
                }
                else
                {
                    uint ticksElapsed;
                    unchecked
                    {
                        ticksElapsed = (uint)Environment.TickCount - startTicks;
                    }

                    double bps = ((double)bytesDownloaded * 8000.0) / (double)ticksElapsed;
                    TimeSpan remain = new TimeSpan(((long)((selectedFilesSize - bytesDownloaded) / (bps/8))) * 10000000L);

                    mediaQueue.ReportStatus($"Downloading file {n + 1} of {queue.Count}. Time remaining: {remain.FmtCustom()} Mbps: {(bps / (1024 * 1024)):#,###.###}");
                }

                string dstFilepath = Path.Combine(sourceConfig.DestinationDirectory, fi.OriginalFilename);
                MediaFile.MakeFilepathUnique(ref dstFilepath);

                FetchFile(fi.Url, dstFilepath);
                bytesDownloaded += fi.Size;
                ++n;

                // Add to the destination queue
                mediaQueue.Add(new ProcessFileInfo(
                    dstFilepath,
                    fi.Size,
                    fi.OriginalFilename,
                    fi.OriginalDateCreated ?? DateTime.MinValue,
                    fi.OriginalDateModified ?? DateTime.MinValue));

                if (fi.BookmarkDate.HasValue && newestSelection < fi.BookmarkDate)
                {
                    newestSelection = fi.BookmarkDate.Value;
                }
            }

            TimeSpan elapsed;
            unchecked
            {
                uint ticksElapsed = (uint)Environment.TickCount - startTicks;
                elapsed = new TimeSpan(ticksElapsed * 10000L);
            }

            mediaQueue.ReportStatus(null);
            mediaQueue.ReportProgress($"Download complete. {queue.Count} files, {bytesDownloaded / (1024.0 * 1024.0): #,##0.0} MB, {elapsed.FmtCustom()} elapsed");
            if (newestSelection > DateTime.MinValue)
            {
                if (sourceConfig.SetBookmark(m_bookmarkPath, newestSelection))
                {
                    mediaQueue.ReportProgress($"Bookmark Set to {newestSelection}");
                }
            }
        }

        static DateTime? GetBookmarkDate(XPathNavigator fileNode)
        {
            // Ideally we would use DateTaken. But for photos, it is in localTime even though
            // the provided metadata uses the 'Z' suffix. For videos it is in UCT. This risks
            // missing future files if the last file is a video and an upcoming file in the next
            // patch is a photo.
            //
            // Follow the same pattern as with MediaFile.GetBookmarkDate
            // 1. Parse from filename.
            // 2. Use DateTaken from photo but treat to local time with no timezone shift.
            // 3. Use DateTaken from video (really DateEncoded) but shift to local time.
            // 4. Use file system DateCreated and shift to local time.
          
            // From filename
            DateTime dt;
            if (MediaFile.TryParseDateTimeFromFilename(fileNode.XPVal("name"), out dt))
            {
                return dt;
            }

            // If JPEG
            if (string.Equals(fileNode.XPVal("file/mimeType"), "image/jpeg", StringComparison.Ordinal)
                && TryParseDt(fileNode.XPVal("photo/takenDateTime"), out dt))
            {
                // Change to local without shifting the time value
                return new DateTime(dt.Ticks, DateTimeKind.Local);
            }

            // If MP4
            if (string.Equals(fileNode.XPVal("file/mimeType"), "video/mp4", StringComparison.Ordinal)
                && TryParseDt(fileNode.XPVal("photo/takenDateTime"), out dt))
            {
                // Change to local WITH shifting the time value
                return dt.ToLocalTime();
            }

            if (TryParseDt(fileNode.XPVal("fileSystemInfo/createdDateTime"), out dt))
            {
                // Change to local WITH shifting the time value
                return dt.ToLocalTime();
            }
            return null;
        }

        static bool TryParseDt(string value, out DateTime result)
        {
            if (value == null)
            {
                result = DateTime.MinValue;
                return false;
            }
            return (DateTime.TryParse(value, CultureInfo.InvariantCulture,
                DateTimeStyles.RoundtripKind | DateTimeStyles.NoCurrentDateDefault, out result));
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

        void FetchFile(string url, string filepath)
        {
            using (var response = FetchResponse(url))
            {
                using (var dst = new FileStream(filepath, FileMode.Create, FileAccess.Write, FileShare.Read))
                {
                    using (var src = response.GetResponseStream())
                    {
                        src.CopyTo(dst);
                    }
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

        class OdFileInfo
        {
            public OdFileInfo()
            {
            }

            public OdFileInfo(XPathNavigator fileNode)
            {
                Url = string.Concat(c_oneDriveItemPrefix, fileNode.XPVal("id"), c_oneDriveContentSuffix);
                OriginalFilename = fileNode.XPVal("name");
                DateTime dt;
                if (TryParseDt(fileNode.XPVal("fileSystemInfo/createdDateTime"), out dt))
                {
                    OriginalDateCreated = dt;
                }
                if (TryParseDt(fileNode.XPVal("fileSystemInfo/lastModifiedDateTime"), out dt))
                {
                    OriginalDateModified = dt;
                }
                BookmarkDate = GetBookmarkDate(fileNode);
                if (!long.TryParse(fileNode.XPVal("size"), out Size))
                {
                    Size = 0L;
                }
            }

            public string Url;
            public string OriginalFilename;
            public DateTime? OriginalDateCreated;
            public DateTime? OriginalDateModified;
            public DateTime? BookmarkDate;
            public long Size;
        }
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
