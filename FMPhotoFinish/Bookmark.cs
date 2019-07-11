using System;
using System.Xml.Linq;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Globalization;
using FileMeta;

namespace FMPhotoFinish
{
    // Because DateTaken is in local time. Bookmarks are also handled in local time.
    class IncrementalBookmark
    {
        static Encoding s_UTF8_NoBOM = new UTF8Encoding(false);

        const DateTimeStyles c_dateParseStyle = DateTimeStyles.AssumeLocal | DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.NoCurrentDateDefault | DateTimeStyles.RoundtripKind;
        const string c_dateFormat = "yyyy'-'MM'-'dd'T'HH':'mm':'ss.FFF";
        const string c_bookmarkFilename = "FMPhotoFinish_Incremental.json";
        const string c_rootEle = "root";
        const string c_bookmarksEle = "bookmarks";
        const string c_itemEle = "item";
        const string c_srcPathEle = "sourcePath";
        const string c_newestEle = "newestSelected";
        const string c_typeAttr = "type";
        const string c_object = "object";
        const string c_array = "array";

        string m_destinationPath;

        public IncrementalBookmark(string destinationPath)
        {
            m_destinationPath = destinationPath;
        }

        public DateTime? GetBookmark(string sourcePath)
        {
            var doc = LoadBookmarks(m_destinationPath);

            XElement bookmark =
                (from el in doc.Element(c_bookmarksEle).Elements(c_itemEle)
                where el.Element(c_srcPathEle).Value.Equals(sourcePath, StringComparison.OrdinalIgnoreCase)
                select el).FirstOrDefault();

            if (bookmark == null) return null;

            DateTime dt;
            if (!DateTime.TryParse(bookmark.Element(c_newestEle).Value,
                CultureInfo.InvariantCulture, c_dateParseStyle, out dt)) return null;

            return dt;
        }

        public void SetBookmark(string sourcePath, DateTime bookmarkDate)
        {
            var doc = LoadBookmarks(m_destinationPath);

            var dateStr = bookmarkDate.ToLocalTime().ToString(c_dateFormat, CultureInfo.InvariantCulture);

            XElement bookmark =
                (from el in doc.Element(c_bookmarksEle).Elements(c_itemEle)
                 where el.Element(c_srcPathEle).Value.Equals(sourcePath, StringComparison.OrdinalIgnoreCase)
                 select el).FirstOrDefault();

            if (bookmark == null)
            {
                bookmark = new XElement(c_itemEle,
                    new XAttribute(c_typeAttr, c_object),
                    new XElement(c_srcPathEle, sourcePath),
                    new XElement(c_newestEle, dateStr));
                doc.Element(c_bookmarksEle).Add(bookmark);
            }
            else
            {
                bookmark.Element("c_newestEle").SetValue(dateStr);
            }

            SaveBookmarks(m_destinationPath, doc);
        }

        private static XElement LoadBookmarks(string destinationPath)
        {
            XElement doc = null;

            string bookmarkPath = Path.Combine(destinationPath, c_bookmarkFilename);
            if (File.Exists(bookmarkPath))
            {
                try
                {
                    using (var stream = new FileStream(bookmarkPath, FileMode.Open, FileAccess.Read, FileShare.Read))
                    {
                        using (var jsonReader = System.Runtime.Serialization.Json.JsonReaderWriterFactory.CreateJsonReader(stream, new System.Xml.XmlDictionaryReaderQuotas()))
                        {
                            doc = XElement.Load(jsonReader);
                        }
                    }
                }
                catch
                {
                    doc = null;
                }
            }

            if (doc == null)
            {
                doc = new XElement(c_rootEle,
                    new XAttribute(c_typeAttr, c_object),
                    new XElement("bookmarks",
                        new XAttribute(c_typeAttr, c_array)));
                var fullDocument = new XDocument(doc);
            }

            DumpXml(doc);

            return doc;
        }

        private static void SaveBookmarks(string destinationPath, XElement doc)
        {
            DumpXml(doc);
            string bookmarkPath = Path.Combine(destinationPath, c_bookmarkFilename);

            using (var stream = new FileStream(bookmarkPath, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                using (var jsonWriter = System.Runtime.Serialization.Json.JsonReaderWriterFactory.CreateJsonWriter(stream, s_UTF8_NoBOM, false, true))
                {
                    doc.WriteTo(jsonWriter);
                }
            }
        }

#if DEBUG
        public static void DumpXml(XElement xml, TextWriter writer = null)
        {
            if (writer == null) writer = Console.Out;

            var settings = new System.Xml.XmlWriterSettings();
            settings.Indent = true;
            settings.CloseOutput = false;
            using (var xmlwriter = System.Xml.XmlWriter.Create(writer, settings))
            {
                xml.WriteTo(xmlwriter);
            }
            writer.WriteLine();
            writer.Flush();
        }
#endif

    }
}
