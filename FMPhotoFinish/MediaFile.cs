using System;
using WinShell;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using FileMeta;

namespace FMPhotoFinish
{
    enum MediaType
    {
        Unsupported = 0, // Unsupported type
        Image = 1,  // .jpg .jpeg
        Video = 2,  // .mp4 .avi .mov .mpg .mpeg
        Audio = 3   // .m4a .mp3 .wav 
    }

    /// <summary>
    /// Performs operations on a media file such as metadata changes, rotation, recoding, etc.
    /// </summary>
    class MediaFile : IDisposable
    {
        const string c_ffmpegVideoSettings = "-pix_fmt yuv420p -c:v libx264 -profile:v main -level:v 3.1 -crf 18";
        const string c_ffmpegAudioSettings = "-c:a aac"; // Let it use the default quality settings

        const string c_jpgExt = ".jpg";
        const string c_mp4Ext = ".mp4";
        const string c_m4aExt = ".m4a";

        const string c_timezoneKey = "timezone";

        #region Static Members

        static Encoding s_Utf8NoBOM = new UTF8Encoding(false);

        static Dictionary<string, MediaType> s_mediaExtensions = new Dictionary<string, MediaType>()
        {
            {c_jpgExt, MediaType.Image},
            {".jpeg", MediaType.Image},
            {c_mp4Ext, MediaType.Video},
            {".avi", MediaType.Video},
            {".mov", MediaType.Video},
            {".mpg", MediaType.Video},
            {".mpeg", MediaType.Video},
            {c_m4aExt, MediaType.Audio},
            {".mp3", MediaType.Audio},
            {".wav", MediaType.Audio}
        };

        static HashSet<string> s_isomExtensions = new HashSet<string>()
        {
            c_mp4Ext, ".mov", c_m4aExt
        };

        // Preferred formats (by media type
        static string[] s_preferredExtensions = new string[]
        {
            null,   // Unknown
            c_jpgExt, // Image
            c_mp4Ext, // Video
            c_m4aExt  // Audio
        };

        public static MediaType GetMediaType(string filenameOrExtension)
        {
            int ext = filenameOrExtension.LastIndexOf('.');
            if (ext >= 0)
            {
                MediaType t;
                if (s_mediaExtensions.TryGetValue(filenameOrExtension.Substring(ext).ToLowerInvariant(), out t))
                {
                    return t;
                }
            }
            return MediaType.Unsupported;
        }

        public static bool IsSupportedMediaType(string filenameOrExtension)
        {
            return GetMediaType(filenameOrExtension) != MediaType.Unsupported;
        }

        public static void MakeFilepathUnique(ref string filepath)
        {
            if (!File.Exists(filepath)) return;

            string basepath = Path.Combine(Path.GetDirectoryName(filepath), Path.GetFileNameWithoutExtension(filepath));
            string extension = Path.GetExtension(filepath);

            int index = 1;

            // Strip and parse any parenthesized index
            int p = basepath.Length;
            if (p > 2 && basepath[p-1] == ')' && char.IsDigit(basepath[p-2]))
            {
                p -= 2;
                while (char.IsDigit(basepath[p - 1])) --p;
                if (p > 1 && basepath[p-1] == '(')
                {
                    --p;
                    // Found a parenthesized index
                    index = int.Parse(basepath.Substring(p + 1, basepath.Length - p - 2)) + 1; // Update the index to the next value
                    basepath = basepath.Substring(0, p); // Remove the parenthesized index from the basepath;
                }
            }

            // Generate a unique filename
            do
            {
                filepath = $"{basepath}({index}){extension}";
                ++index;
            } while (File.Exists(filepath));
        }

        /// <summary>
        /// Parses a DateTime from a file naming convention.
        /// </summary>
        /// <param name="filename">A Filename (without path).</param>
        /// <param name="result">A parsed dateTime (if successful)</param>
        /// <returns>True if successful, else, false</returns>
        /// <remarks>
        /// Most phones give the photo a filename that is derived from the date and time
        /// that the photo or video was taken (within a few seconds). This function detects
        /// whether it's a supported naming convention and parses the date and time.
        /// </remarks>
        public static bool TryParseDateTimeFromFilename(string filename, out DateTime result)
        {
            result = DateTime.MinValue;
            var sb = new StringBuilder();
            foreach(char c in filename)
            {
                if (char.IsDigit(c)) sb.Append(c);
            }
            var digits = sb.ToString();
            if (digits.Length < 14) return false;

            int year = int.Parse(digits.Substring(0, 4));
            int month = int.Parse(digits.Substring(4, 2));
            int day = int.Parse(digits.Substring(6, 2));
            int hour = int.Parse(digits.Substring(8, 2));
            int minute = int.Parse(digits.Substring(10, 2));
            int second = int.Parse(digits.Substring(12, 2));
            int millisecond = 0;
            if (digits.Length > 14)
            {
                millisecond = (int)(int.Parse(digits.Substring(14))
                    * Math.Pow(10, 17 - digits.Length));
            }
            if (year < 1900 || year > 2200) return false;
            if (month < 1 || month > 12) return false;
            if (day < 1 || day > 31) return false;
            if (hour < 0 || hour > 23) return false;
            if (minute < 0 || minute > 59) return false;
            if (second < 0 || second > 59) return false;

            try
            {
                result = new DateTime(year, month, day, hour, minute, second, millisecond, DateTimeKind.Local);
            }
            catch (Exception)
            {
                return false; // Probaby a month with too many days.
            }

            return true;
        }

        /// Number of seconds of tolerance when comparing for hours offset
        static long c_absoluteTolerance = (new TimeSpan(24, 0, 0)).Ticks;
        static long c_fractionalTolerance = (new TimeSpan(0, 0, 60)).Ticks;
        static long c_ticksThirtyMinutes = (new TimeSpan(0, 30, 0)).Ticks;

        public bool TryGetTimezoneOffset(DateTime a, DateTime b, out int offset)
        {
            offset = 0;
            TimeSpan ts = a.Subtract(b);

            if (Math.Abs(ts.Ticks) >= c_absoluteTolerance) return false;

            // Round to nearest half-hour
            TimeSpan tsRounded;
            {
                long rd = ts.Ticks;
                rd += (rd > 0) ? (c_ticksThirtyMinutes/2) : -(c_ticksThirtyMinutes/2);
                rd = (rd / c_ticksThirtyMinutes) * c_ticksThirtyMinutes;
                tsRounded = new TimeSpan(rd);
            }

            // Check whether the remainder exceeds the fractional tolerance
            if (Math.Abs(ts.Ticks - tsRounded.Ticks) > c_fractionalTolerance) return false;

#if DEBUG
            Debug.WriteLine($"   Offset: {tsRounded.TotalHours}");
            Debug.WriteLine($"   Tolerance: {(new TimeSpan(ts.Ticks - tsRounded.Ticks)).TotalSeconds:F3}");
#endif

            offset = (int)tsRounded.TotalMinutes;
            return true;
        }

        // TODO: Set more carefully defined standards for metatag.
        // Key should follow standards for a hashtag. Value should use
        // percent encoding for disallowed characters, underscore for space.
        // So far, disallowed characters are whitespace and &.
        static Regex s_rxMetatag = new Regex(
            @"^&([A-Za-z0-9]+)=([^ \t\r\n&]+)$",
            RegexOptions.CultureInvariant);

        public static bool TryParseMetatag(string s, out string key, out string value)
        {
            var match = s_rxMetatag.Match(s);
            if (!match.Success)
            {
                key = null;
                value = null;
                return false;
            }

            key = match.Groups[1].Value;
            value = MetatagDecode(match.Groups[2].Value);
            return true;
        }

        public static string FormatMetatag(string key, string value)
        {
            return $"&{key}={MetatagEncode(value)}";
        }

        public static string MetatagEncode(string s)
        {
            var sb = new StringBuilder();
            foreach(char c in s)
            {
                switch (c)
                {
                    case ' ':
                        sb.Append('_');
                        break;

                    case '_':
                    case '%':
                    case '&':
                    case '\r':
                    case '\n':
                    case '\t':
                        sb.Append($"%{((int)c):x2}");
                        break;

                    default:
                        sb.Append(c);
                        break;
                }
            }
            return sb.ToString();
        }

        public static string MetatagDecode(string s)
        {
            var sb = new StringBuilder();
            for (int i=0; i<s.Length; ++i)
            {
                char c = s[i];
                if (c == '_')
                {
                    sb.Append(' ');
                }
                else if (c == '%' && i < s.Length+2)
                {
                    int n;
                    if (int.TryParse(s.Substring(i + 1, 2), System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture,
                        out n) && n > 0 && n < 256)
                    {
                        sb.Append((char)n);
                        i += 2;
                    }
                    else
                    {
                        sb.Append('%');
                    }
                }
                else
                {
                    sb.Append(c);
                }
            }
            return sb.ToString();
        }

        #endregion

        #region Delegates

        public delegate void ProgressReporter(string progress);

#endregion

        string m_filepath;
        string m_originalFilename;
        MediaType m_mediaType;

        Dictionary<PROPERTYKEY, object> m_propsToSet = new Dictionary<PROPERTYKEY, object>();

        // File System Values (In local time)
        DateTime m_fsDateCreated;
        DateTime m_fsDateModified;

        // Critical values from the Windows Property System
        DateTime? m_psItemDate;
        TimeSpan? m_psDuration;

        // Metatag values from Property System Keywords
        TimeZoneTag m_mtTimezone;

        // Values from ExifTool
        TimeZoneTag m_etTimezone;

        // Master Values
        bool updateMetadata = false; // Metadata should be updated upon disposal
        DateTime? m_creationDate;
        TimeZoneTag m_timezone;
        string m_make;
        string m_model;

        public MediaFile(string filepath, string originalFilename)
        {
            m_filepath = filepath;
            m_originalFilename = originalFilename ?? Path.GetFileName(filepath);
            string ext = Path.GetExtension(filepath).ToLowerInvariant();
            if (!s_mediaExtensions.TryGetValue(ext, out m_mediaType))
            {
                throw new InvalidOperationException($"Media type '{ext}' is not supported.");
            }

            Orientation = 1; // Defaults to normal/vertical

            // Load Windows Property Store properties
            using (var propstore = PropertyStore.Open(filepath, false))
            {
                m_fsDateCreated = (DateTime)propstore.GetValue(PropertyKeys.DateCreated);
                m_fsDateModified = (DateTime)propstore.GetValue(PropertyKeys.DateModified);
                m_psItemDate = (DateTime?)propstore.GetValue(PropertyKeys.DateTaken);
                if (!m_psItemDate.HasValue)
                    m_psItemDate = (DateTime?)propstore.GetValue(PropertyKeys.DateEncoded);

                {
                    var duration = propstore.GetValue(PropertyKeys.Duration);
                    if (duration != null)
                        m_psDuration = new TimeSpan((long)(ulong)duration);
                }
                Orientation = (int)(ushort)(propstore.GetValue(PropertyKeys.Orientation) ?? (ushort)1);

                // Todo: Add Make and Model here


                // Keywords may be used to store custom metadata
                var keywords = (string[])propstore.GetValue(PropertyKeys.Keywords);
                if (keywords != null)
                {
                    foreach(string s in keywords)
                    {
                        string key, value;
                        TimeZoneTag timezone;
                        if (TryParseMetatag(s, out key, out value)
                            && key.Equals(c_timezoneKey, StringComparison.OrdinalIgnoreCase)
                            &&  TimeZoneTag.TryParse(value, out timezone))
                        {
                            m_mtTimezone = timezone;
                        }
                    }
                }
            }

            // Load ExifTool Properties
            {
                var exifProperties = new List<KeyValuePair<string, string>>();
                s_exifTool.GetProperties(m_filepath, exifProperties);
                string software = null;

                foreach (var pair in exifProperties)
                {
                    int colon = pair.Key.IndexOf(':');
                    string key = (colon >= 0) ? pair.Key.Substring(colon + 1) : pair.Key;
                    switch (key.ToLowerInvariant())
                    {
                        case "timezone":
                            {
                                TimeZoneTag tz;
                                if (TimeZoneTag.TryParse(pair.Value, out tz))
                                {
                                    m_etTimezone = tz;
                                }
                            }
                            break;

                        case "make":
                            if (m_make == null)
                            {
                                m_make = pair.Value;
                            }
                            break;

                        case "model":
                            if (m_make == null)
                            {
                                m_model = pair.Value;
                            }
                            break;

                        // For AVI Video
                        case "software":
                            software = pair.Value;
                            break;
                    }
                }

                if (m_make == null) m_make = software;
                if (m_model == null) m_model = software;
            }

        }

        public string Filepath { get { return m_filepath; } }

        public MediaType MediaType { get { return m_mediaType; } }

        public DateTime? OriginalDateCreated { get; set; }
        public DateTime? OriginalDateModified { get; set; }
        public int Orientation { get; private set; }

        public TimeZoneTag Timezone
        {
            get { return m_timezone ?? m_mtTimezone ?? m_etTimezone ?? TimeZoneTag.Unknown; }
            set
            {
                m_timezone = value;
                TimezoneSource = "Explicit";
            }
        }

        public string TimezoneSource { get; private set; }
        
        public void RotateToVertical()
        {
            JpegRotator.RotateToVertical(m_filepath);
            Orientation = 1; // Normal
        }

        public string PreferredFormat { get { return s_preferredExtensions[(int)m_mediaType]; } }

        public bool IsPreferredFormat { get { return string.Equals(Path.GetExtension(m_filepath), PreferredFormat, StringComparison.OrdinalIgnoreCase); } }

        public bool TranscodeToPreferredFormat(ProgressReporter reporter)
        {
            string ext = Path.GetExtension(m_filepath);
            if (string.Equals(ext, PreferredFormat, StringComparison.OrdinalIgnoreCase))
                return true;

            if (m_mediaType == MediaType.Image)
            {
                Debug.Assert(ext.Equals(".jpeg", StringComparison.OrdinalIgnoreCase));
                ChangeExtensionTo(c_jpgExt);
                return true;
            }

            if (m_mediaType != MediaType.Audio && m_mediaType != MediaType.Video)
                return false;

            return Transcode(reporter);
        }

        /// <summary>
        /// Attempt to determine the timezone of media files - especially those
        /// with UTC metadata. (See remarks for detail.) 
        /// </summary>
        /// <remarks>
        /// <para>Video and audio files in ISOM-derived formats (.mov, .mp4, .m4a, etc.) and
        /// video in .avi format store the date/time of the event in UTC. Meanwhile, photos
        /// in .jpg store the event date/time in local time. In order to render the times
        /// of mixed video and photo content consistently we need to know the timezone of the
        /// UTC items. Knowing the timezone of .jpg (local time) items is nice but not as
        /// important.
        /// </para>
        /// <para>This method attempts to determine the timezone through various methods and
        /// stores it in the custom &datetime metatag. Some of the methods it uses only work
        /// for formats with DateTime in UTC - which is convenient since that's the more
        /// important case.
        /// </para>
        /// <para>For cameras without a timezone setting, a local time is often stored in
        /// a UTC field. In that case, the timezone is set to "+00:00".
        /// </para>
        /// </remarks>
        public bool DetermineTimezone()
        {
            // DateTime is not known so timezone is irrelevant
            if (!m_psItemDate.HasValue) return false;

            if (!TimeZoneTag.IsNullOrUnknown(m_mtTimezone))
            {
                m_timezone = m_mtTimezone;
                TimezoneSource = "Existing";
                updateMetadata = true;
                return true;
            }

            if (!TimeZoneTag.IsNullOrUnknown(m_etTimezone))
            {
                m_timezone = m_etTimezone;
                TimezoneSource = "MakerNote";
                updateMetadata = true;
                return true;
            }

            /* Video or audio files from phones generally store the ItemTime in UTC, but they
             * name the file according to local time. So, we can detect the timezone
             * by comparing the ItemTime against the date/time parsed from the filename.
             * Photos from phones use the same naming convention but store local time so
             * the calculation should result in a zero offset. It means we can't determine
             * the timezone for photos but we have the local time which is more important.*/
            int timezone;
            DateTime dt;
            if (m_psItemDate.Value.Kind == DateTimeKind.Utc
                && TryParseDateTimeFromFilename(m_originalFilename, out dt)
                && TryCalcTimezoneFromMatch(dt, out timezone))
            {
                m_timezone = new TimeZoneTag(timezone, TimeZoneKind.Normal);
                TimezoneSource = "Filename";
                updateMetadata = true;
                return true;
            }

            /* Compare with the file system timestamp in local time. If the media is being read
             * from FAT (which stores in local time) then a zero offset means the camera uses
             * local time in the UTC field.
             */
            if (m_psItemDate.Value.Kind == DateTimeKind.Utc)
            {
                if (TryCalcTimezoneFromMatch((OriginalDateCreated ?? m_fsDateCreated).ToLocalTime(), out timezone)
                    && timezone == 0)
                {
                    m_timezone = TimeZoneTag.ForceLocal;
                    TimezoneSource = "FileSystem";
                    updateMetadata = true;
                    return true;
                }

                if (TryCalcTimezoneFromMatch((OriginalDateModified ?? m_fsDateModified).ToLocalTime(), out timezone)
                    && timezone == 0)
                {
                    m_timezone = TimeZoneTag.ForceLocal;
                    TimezoneSource = "FileSystem";
                    updateMetadata = true;
                    return true;
                }
            }

            return false;
        }



        #region IDisposable Support

        protected virtual void Dispose(bool disposing)
        {
            if (m_filepath == null) return;

            if (disposing)
            {
                if (m_propsToSet.Count > 0)
                {
#if DEBUG
                    Debug.WriteLine(Path.GetFileName(m_filepath));
                    foreach (var pair in m_propsToSet)
                    {
                        Debug.WriteLine($"   {pair.Value}");
                    }
#endif
                    using (var ps = PropertyStore.Open(m_filepath, true))
                    {
                        foreach (var pair in m_propsToSet)
                        {
                            ps.SetValue(pair.Key, pair.Value);
                        }
                        ps.Commit();
                    }
                }
            }

            else
            {
                System.Diagnostics.Debug.Fail("Failed to dispose of MediaFile.");
            }
        }

#if DEBUG
        ~MediaFile()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(false);
        }
#endif

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
#if DEBUG
            GC.SuppressFinalize(this);
#endif
        }

#endregion // IDisposable Support

#region Private Members

        bool ChangeExtensionTo(string newExt)
        {
            bool result = true;
            try
            {
                string newPath = Path.ChangeExtension(m_filepath, newExt);
                MakeFilepathUnique(ref newPath);
                File.Move(m_filepath, newPath);
                m_filepath = newPath;
            }
            catch
            {
                result = false;
            }
            return result;
        }

        const string c_FFMpeg = "FFMpeg.exe";
               
        bool Transcode(ProgressReporter reporter)
        {
            // If inbound file does not have a duration, it's not a transcodeable media file
            if (m_psDuration == null)
            {
                return false;
            }

            string newPath = Path.ChangeExtension(m_filepath, PreferredFormat);
            MakeFilepathUnique(ref newPath);

            Process transcoder = null;
            bool result = false;
            try
            {
                // Compose arguments
                string arguments;
                if (m_mediaType == MediaType.Video)
                {
                    arguments = $"-hide_banner -i {m_filepath} {c_ffmpegVideoSettings} {c_ffmpegAudioSettings} {newPath}";
                }
                else if (m_mediaType == MediaType.Audio)
                {
                    arguments = $"-hide_banner -i {m_filepath} {c_ffmpegAudioSettings} {newPath}";
                }
                else
                {
                    throw new InvalidOperationException();
                }

                // Prepare process start
                var psi = new ProcessStartInfo(c_FFMpeg, arguments);
                psi.UseShellExecute = false;
                psi.CreateNoWindow = true; // Set to false if you want to monitor
                psi.RedirectStandardError = true;
                psi.StandardErrorEncoding = s_Utf8NoBOM;

                transcoder = Process.Start(psi);

                bool wroteProgress = false;
                var sb = new StringBuilder();
                for (; ; )
                {
                    int i = transcoder.StandardError.Read();
                    if (i < 0) break;
                    if (i == '\n')
                    {
                        sb.Clear();
                    }
                    else if (i == '\r' && transcoder.StandardError.Peek() != '\n')
                    {
                        reporter?.Invoke("Transcoding: " + sb.ToString());
                        wroteProgress = true;
                        sb.Clear();
                    }
                    else
                    {
                        sb.Append((char)i);
                    }
                }
                transcoder.WaitForExit();
                if (wroteProgress)
                {
                    reporter?.Invoke(null);
                }

                result = transcoder.ExitCode == 0;
            }
            finally
            {
                if (transcoder != null)
                {
                    transcoder.Dispose();
                    transcoder = null;
                }
            }

            if (result)
            {
                // Confirm the transcoding by reading out the duration and comparing with original (if available)
                var isom = FileMeta.IsomCoreMetadata.TryOpen(newPath);
                if (isom == null)
                {
                    result = false;
                }
                else
                {
                    using (isom)
                    {
                        Debug.Assert(m_psDuration != null); // Should have exited early if duration is null.
                        if (isom.Duration == null
                            || Math.Abs(m_psDuration.Value.Ticks - isom.Duration.Ticks) > (250L * 10000L)) // 1/4 second
                        {
                            result = false;
                        }
                    }
                }
            }

            // If successful, replace original with transcoded. If failed, delete the transcoded version.
            if (result)
            {
                File.Delete(m_filepath);
                m_filepath = newPath;
                if (m_make != null)
                    m_propsToSet.Add(PropertyKeys.Make, m_make);
                if (m_model != null)
                    m_propsToSet.Add(PropertyKeys.Model, m_model);
            }
            else
            {
                File.Delete(newPath);
            }

            return result;
        }

        bool TryCalcTimezoneFromMatch(DateTime timeToMatch, out int timezone)
        {
            Debug.Assert(m_psItemDate.HasValue); // Caller should check for value first.
            Debug.Assert(m_psItemDate.Value.Kind == DateTimeKind.Local
                || m_psItemDate.Value.Kind == DateTimeKind.Utc);

            // Check if the match time is offset by a whole number of hours (given the tolerance of TryGetHourOffset)
            if (TryGetTimezoneOffset(m_psItemDate.Value, timeToMatch, out timezone)) return true;

            // If the media has a duration, compare against both the start and end.
            // Certain timestamp events correspond to one or the other and without
            // incorporating duration, the tolerance might not be satisfied.
            if (m_psDuration.HasValue)
            {
                if (TryGetTimezoneOffset(m_psItemDate.Value.Add(m_psDuration.Value), timeToMatch, out timezone)) return true;
                if (TryGetTimezoneOffset(m_psItemDate.Value.Subtract(m_psDuration.Value), timeToMatch, out timezone)) return true;
            }
            return false;
        }

#endregion // Private Members

#region Private Static Members

        // Cache whether a propery is copyable
        static Dictionary<PROPERTYKEY, bool> s_propertyIsCopyable = new Dictionary<PROPERTYKEY, bool>();

        static bool IsCopyable(PROPERTYKEY pk)
        {
            bool result;
            if (s_propertyIsCopyable.TryGetValue(pk, out result))
            {
                return result;
            }

            var desc = s_propSystem.GetPropertyDescription(pk);
            result = desc != null
                && desc.ValueTypeIsSupported
                && (desc.TypeFlags & PROPDESC_TYPE_FLAGS.PDTF_ISINNATE) == 0;
            s_propertyIsCopyable[pk] = result;

            return result;
        }

#endregion

#region Init and Shutdown

        static PropertySystem s_propSystem = new PropertySystem();
        static ExifToolWrapper.ExifTool s_exifTool = new ExifToolWrapper.ExifTool();
        static readonly StaticDisposer s_psDisposer = new StaticDisposer();

        private sealed class StaticDisposer
        {
            ~StaticDisposer()
            {
                if (s_propSystem != null)
                {
                    s_propSystem.Dispose();
                    s_propSystem = null;
                }
                if (s_exifTool != null)
                {
                    s_exifTool.Dispose();
                    s_exifTool = null;
                }
            }
        }

#endregion
    }

    /// <summary>
    /// Properties defined here: https://msdn.microsoft.com/en-us/library/windows/desktop/dd561977(v=vs.85).aspx
    /// </summary>
    static class PropertyKeys
    {
        public static PROPERTYKEY ItemDate = new PROPERTYKEY("f7db74b4-4287-4103-afba-f1b13dcd75cf", 100); // System.ItemDate
        public static PROPERTYKEY DateEncoded = new PROPERTYKEY("2e4b640d-5019-46d8-8881-55414cc5caa0", 100); // System.Media.DateEncoded
        public static PROPERTYKEY DateTaken = new PROPERTYKEY("14b81da1-0135-4d31-96d9-6cbfc9671a99", 36867); // System.Photo.DateTaken
        public static PROPERTYKEY DateCreated = new PROPERTYKEY("b725f130-47ef-101a-a5f1-02608c9eebac", 15); // System.DateCreated (from file system)
        public static PROPERTYKEY DateModified = new PROPERTYKEY("b725f130-47ef-101a-a5f1-02608c9eebac", 14); // System.DateModified (from file system)
        public static PROPERTYKEY Keywords = new PROPERTYKEY("f29f85e0-4ff9-1068-ab91-08002b27b3d9", 5); // System.Keywords (tags)
        public static PROPERTYKEY Orientation = new PROPERTYKEY("14B81DA1-0135-4D31-96D9-6CBFC9671A99", 274);
        public static PROPERTYKEY Duration = new PROPERTYKEY("64440490-4C8B-11D1-8B70-080036B11A03", 3);
        public static PROPERTYKEY Make = new PROPERTYKEY("14b81da1-0135-4d31-96d9-6cbfc9671a99", 271);
        public static PROPERTYKEY Model = new PROPERTYKEY("14b81da1-0135-4d31-96d9-6cbfc9671a99", 272);
    }

}
