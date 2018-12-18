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
        const string c_makeKey = "make";
        const string c_modelKey = "model";
        const string c_originalFilenameKey = "originalFilename";

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

        // Renaming patterns to make sure that image names appear in the order they were taken
        static readonly RenamePattern[] s_renamePatterns = new RenamePattern[]
        {
            new RenamePattern(@"^MVI_(\d{4}).AVI$", @"IMG_$1.AVI"), // Older Canon Cameras
            new RenamePattern(@"^SND_(\d{4}).WAV$", @"IMG_$1.WAV"), // Older Canon Cameras with Voice Annotation Feature
            new RenamePattern(@"^MVI_(\d{4}).MP4$", @"IMG_$1.MP4"), // Newer Canon Cameras
            new RenamePattern(@"^VID_(\d{8}_\d{6,9}).(mp4|MP4)$", @"IMG_$1.$2") // Android Phones
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

        public static bool TryGetTimezoneOffset(DateTime a, DateTime b, out int offset)
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

            offset = (int)tsRounded.TotalMinutes;
            return true;
        }

        #endregion

        #region Delegates

        public delegate void ProgressReporter(string progress);

#endregion

        string m_filepath;
        MediaType m_mediaType;
        string m_originalFilename;      // Filename when process started.

        // File System Values (In local time)
        DateTime m_fsDateCreated;
        DateTime m_fsDateModified;

        // Values from the Windows Property System
        DateTime? m_psDateTaken;
        DateTime? m_psDateEncoded;
        TimeSpan? m_psDuration;

        // Metatag values from Property System Keywords
        TimeZoneTag m_mtTimezone;
        string m_mtOriginalFilename;    // Historically original filename.

        // Valies from IsomCoreMetadata
        DateTime? m_isomCreationTime;

        // Values from ExifTool
        DateTime? m_etDateTimeOriginal; // EXIF:DateTimeOriginal (.jpg, .jpeg) - RIFF:DateTimeOriginal (.avi) - in local time.
        TimeZoneTag m_etTimezone;

        // Master Values (to be preserved and set)
        bool m_updateMetadata = false; // Metadata should be updated upon disposal
        DateTime? m_creationDate;
        TimeZoneTag m_timezone;
        string m_make;
        string m_model;

        public MediaFile(string filepath, string originalFilename)
        {
            m_filepath = filepath;
            if (originalFilename == null)
            {
                m_originalFilename = Path.GetFileName(filepath);
            }
            else
            {
                m_originalFilename = originalFilename;
                if (!string.Equals(originalFilename, Path.GetFileName(filepath)))
                {
                    m_updateMetadata = true;    // Need to store originalFilename
                }
            }
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
                Debug.Assert(m_fsDateCreated.Kind == DateTimeKind.Utc);
                m_fsDateModified = (DateTime)propstore.GetValue(PropertyKeys.DateModified);
                Debug.Assert(m_fsDateModified.Kind == DateTimeKind.Utc);
                m_psDateTaken = (DateTime?)propstore.GetValue(PropertyKeys.DateTaken);
                Debug.Assert(!m_psDateTaken.HasValue || m_psDateTaken.Value.Kind == DateTimeKind.Local);
                m_psDateEncoded = (DateTime?)propstore.GetValue(PropertyKeys.DateEncoded);
                {
                    var duration = propstore.GetValue(PropertyKeys.Duration);
                    if (duration != null)
                        m_psDuration = new TimeSpan((long)(ulong)duration);
                }
                Orientation = (int)(ushort)(propstore.GetValue(PropertyKeys.Orientation) ?? (ushort)1);
                m_make = (string)propstore.GetValue(PropertyKeys.Make);
                m_model = (string)propstore.GetValue(PropertyKeys.Model);

                // Keywords may be used to store custom metadata
                var metaTagSet = new MetaTagSet();
                metaTagSet.LoadKeywords((string[])propstore.GetValue(PropertyKeys.Keywords));
                {
                    string value;
                    TimeZoneTag tz;
                    if (metaTagSet.MetaTags.TryGetValue(c_timezoneKey, out value)
                        && TimeZoneTag.TryParse(value, out tz))
                    {
                        m_mtTimezone = tz;
                    }

                    if (metaTagSet.MetaTags.TryGetValue(c_originalFilenameKey, out value))
                    {
                        m_mtOriginalFilename = value;
                    }
                    else
                    {
                        m_updateMetadata = true;    // Need to update originalFilename
                    }
                }
            }

            // Load Isom Property
            if (s_isomExtensions.Contains(ext))
            {
                var isom = FileMeta.IsomCoreMetadata.TryOpen(filepath);
                if (isom != null)
                {
                    using (isom)
                    {
                        m_isomCreationTime = isom.CreationTime;
                        Debug.Assert(!m_isomCreationTime.HasValue || m_isomCreationTime.Value.Kind == DateTimeKind.Utc);
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
                        case "datetimeoriginal": // EXIF from JPEG and Canon .mp4 files - in local time
                            {
                                Debug.Assert(pair.Key.Equals("ExifIFD:DateTimeOriginal") || pair.Key.Equals("RIFF:DateTimeOriginal"));
                                DateTime dt;
                                if (ExifToolWrapper.ExifTool.TryParseDate(pair.Value, DateTimeKind.Local, out dt)) m_etDateTimeOriginal = dt;
                            }
                            break;

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

        public DateTime CreationDate
        {
            get { return m_creationDate ?? m_psDateTaken ?? m_psDateEncoded ?? m_isomCreationTime ?? m_etDateTimeOriginal ?? m_fsDateCreated; }
            set
            {
                m_creationDate = value;
                CreationDateSource = "Explicit";
            }
        }

        public string CreationDateSource { get; private set; }

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

        /// <summary>
        /// Change the filename, if necessary, so that video and audio files are listed in-order
        /// with photos,
        /// </summary>
        /// <returns>True if a rename was necessary. Else, false.</returns>
        /// <remarks>
        /// <para>Some cameras use different filename prefixes for video files vs. photos. When
        /// playing back in filename order, this results in videos being shown out-of-order from
        /// the photos. This method renames files according to certain patterns to ensure that
        /// all are shown in the expected order.
        /// </para>
        /// </remarks>
        public bool SetOrderedName()
        {
            string newName;
            if (RenamePattern.TryGetNewName(s_renamePatterns, m_originalFilename, out newName))
            {
                // Create the new path and make it unique
                string newPath = Path.Combine(Path.GetDirectoryName(m_filepath), newName);
                MakeFilepathUnique(ref newPath);

                // Rename
                File.Move(m_filepath, newPath);
                m_filepath = newPath;
                m_updateMetadata = true;    // Need to store originalFilename
                return true;
            }

            return false;
        }

        /// <summary>
        /// Attempt to determine the date of media files - that is, the date that the event
        /// was recorded. This is the DateTimeOriginal from JPEG EXIF and the creation_date
        /// from .MOV and .MP4 files.
        /// </summary>
        /// <returns>True if the date was successfully determined. Otherwise false.</returns>
        public bool DeterimineCreationDate()
        {
            // We have lots of dates to work from, take them in priority order
            // Note that the reason they are named inconsistently is to match, as closely as reasonable, to the
            // field names in the originating specificaitons or sources.
            if (m_creationDate.HasValue) return true; // Already determined
            if (m_psDateTaken.HasValue)
            {
                m_creationDate = m_psDateTaken;
                CreationDateSource = "Photo.DateTaken";
                // Not necessary to update metadata because this is a primary source
                return true;
            }
            if (m_psDateEncoded.HasValue)
            {
                m_creationDate = m_psDateEncoded;
                CreationDateSource = "Media.DateEncoded";
                // Not necessary to update metadata because this is a primary source
                return true;
            }
            if (m_isomCreationTime.HasValue)
            {
                m_creationDate = m_isomCreationTime;
                CreationDateSource = "Isom.CreationTime";
                m_updateMetadata = true;
                return true;
            }
            if (m_etDateTimeOriginal.HasValue)
            {
                m_creationDate = m_etDateTimeOriginal;
                CreationDateSource = "Exif.DateTimeOriginal";
                m_updateMetadata = true;
                return true;
            }
            // From file naming convention
            {
                DateTime dt;
                if (TryParseDateTimeFromFilename(m_originalFilename, out dt))
                {
                    m_creationDate = dt;
                    CreationDateSource = "Filename";
                    m_updateMetadata = true;
                    return true;
                }
            }
            return false;
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
            if (!m_creationDate.HasValue) return false;

            // Value has already been determined
            if (m_timezone != null) return true;

            if (!TimeZoneTag.IsNullOrUnknown(m_mtTimezone))
            {
                m_timezone = m_mtTimezone;
                TimezoneSource = "MetaTag";
                // Not necessary to update metadata because this is the primary source
                return true;
            }

            if (!TimeZoneTag.IsNullOrUnknown(m_etTimezone))
            {
                m_timezone = m_etTimezone;
                TimezoneSource = "MakerNote";
                m_updateMetadata = true;
                return true;
            }

            /* Video and audio files in ISOM format may have the time stored in the ISOM header
             * field in UTC and in an EXIF field in local time. If so, we can determine the
             * timezone by comparing the two.
             * If the offset is zero then we cannot tell whether timezone is GMT (Britain) or
             * if the camera doesn't have timezone information so we assume the latter because
             * it's more likely. */
            int tzMinutes;
            if (m_isomCreationTime.HasValue && m_etDateTimeOriginal.HasValue
                && TryCalcTimezoneFromMatch(m_isomCreationTime.Value, m_etDateTimeOriginal.Value, m_psDuration, out tzMinutes))
            {
                m_timezone = new TimeZoneTag(tzMinutes, tzMinutes == 0 ? TimeZoneKind.ForceLocal : TimeZoneKind.Normal);
                TimezoneSource = "IsomVsExif";
                m_updateMetadata = true;
                return true;
            }

            /* Phones often name the file according to the local date and time it was taken.
             * This can be compared with available Utc time data isomCreationTime to discover
             * the timezone. */
            DateTime dt;
            if (m_isomCreationTime.HasValue
                && TryParseDateTimeFromFilename(m_originalFilename, out dt)
                && TryCalcTimezoneFromMatch(m_isomCreationTime.Value, dt, m_psDuration, out tzMinutes))
            {
                m_timezone = new TimeZoneTag(tzMinutes, tzMinutes == 0 ? TimeZoneKind.ForceLocal : TimeZoneKind.Normal);
                TimezoneSource = "Filename";
                m_updateMetadata = true;
                return true;
            }

            /* Compare with the file system timestamp in local time. If the media is being read
             * from FAT (which stores in local time) then a zero offset means the camera uses
             * local time in the UTC field. Values other than zero are not trustworthy because
             * the timezone on the computer may have been changed.
             */
            if (m_isomCreationTime.HasValue)
            {
                if (TryCalcTimezoneFromMatch(m_isomCreationTime.Value,
                    (OriginalDateCreated ?? m_fsDateCreated).ToLocalTime(), m_psDuration, out tzMinutes)
                    && tzMinutes == 0)
                {
                    m_timezone = TimeZoneTag.ForceLocal;
                    TimezoneSource = "FileSystem";
                    m_updateMetadata = true;
                    return true;
                }

                if (TryCalcTimezoneFromMatch(m_isomCreationTime.Value,
                    (OriginalDateModified ?? m_fsDateModified).ToLocalTime(), m_psDuration, out tzMinutes)
                    && tzMinutes == 0)
                {
                    m_timezone = TimeZoneTag.ForceLocal;
                    TimezoneSource = "FileSystem";
                    m_updateMetadata = true;
                    return true;
                }
            }

            /* If the creationDate determined earlier is in local timezone then set it to ForceLocal
             */
            if (m_creationDate.HasValue && m_creationDate.Value.Kind == DateTimeKind.Local)
            {
                m_timezone = TimeZoneTag.ForceLocal;
                TimezoneSource = "LocalDate";
                m_updateMetadata = true;
                return true;
            }

            return false;
        }

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
                m_updateMetadata = true; // Need to update originalFilename
                return true;
            }

            if (m_mediaType != MediaType.Audio && m_mediaType != MediaType.Video)
                return false;

            return Transcode(reporter);
        }


        #region IDisposable Support

        protected virtual void Dispose(bool disposing)
        {
            if (m_filepath == null) return;

            if (disposing)
            {
                if (m_updateMetadata)
                {
                    if (m_mediaType == MediaType.Unsupported)
                        throw new ApplicationException("Cannot update metadata on unsupported media type.");

                    // If audio or video, attempt to use Isom to update creationDate
                    bool creationDateStoredByIsom = false;
                    if (m_creationDate.HasValue && (m_mediaType == MediaType.Video || m_mediaType == MediaType.Audio))
                    {
                        var isom = IsomCoreMetadata.TryOpen(m_filepath, true);
                        if (isom != null)
                        {
                            using (isom)
                            {
                                // Convert to UTC (this does nothing if it is already UTC.
                                var dt = (m_timezone != null) ? m_timezone.ToUtc(m_creationDate.Value) : m_creationDate.Value.ToUniversalTime();
                                isom.CreationTime = dt;
                                isom.ModificationTime = dt;
                                isom.Commit();
                            }
                            creationDateStoredByIsom = true;
                        }
                    }

                    using (var ps = PropertyStore.Open(m_filepath, true))
                    {
                        // Prep the metatags with existing values
                        var metaTagSet = new MetaTagSet();
                        metaTagSet.LoadKeywords((string[])ps.GetValue(PropertyKeys.Keywords));

                        // Handle type-specific metadata
                        if (m_mediaType == MediaType.Image)
                        {

                            if (m_creationDate.HasValue && !creationDateStoredByIsom)
                            {
                                // Convert to local (this does nothing if it is already Local.
                                var dt = (m_timezone != null) ? m_timezone.ToLocal(m_creationDate.Value) : m_creationDate.Value.ToLocalTime();
                                ps.SetValue(PropertyKeys.DateTaken, dt);
                            }
                        }

                        // Audio and video both use Isom file format (.mp4 and .m4a)
                        else
                        {
                            if (m_creationDate.HasValue && !creationDateStoredByIsom)
                            {
                                // Convert to UTC (this does nothing if it is already UTC.
                                var dt = (m_timezone != null) ? m_timezone.ToUtc(m_creationDate.Value) : m_creationDate.Value.ToUniversalTime();
                                ps.SetValue(PropertyKeys.DateEncoded, dt);
                            }
                        }

                        if (m_timezone != null)
                            metaTagSet.MetaTags[c_timezoneKey] = m_timezone.ToString();
                        if (!string.IsNullOrEmpty(m_make))
                            ps.SetValue(PropertyKeys.Make, m_make);
                        if (!string.IsNullOrEmpty(m_model))
                            ps.SetValue(PropertyKeys.Model, m_model);

                        // Original filename. If the metatag value exists, it's an historical original
                        // filename from a previous run of this or some other app. If it doesn't exist
                        // then we use the value from the beginning of this job.
                        metaTagSet.MetaTags[c_originalFilenameKey] = m_mtOriginalFilename ?? m_originalFilename;

                        ps.SetValue(PropertyKeys.Keywords, metaTagSet.ToKeywords());

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
                m_updateMetadata = true;
            }
            else
            {
                File.Delete(newPath);
            }

            return result;
        }

        static bool TryCalcTimezoneFromMatch(DateTime utcDate, DateTime localDate, TimeSpan? duration, out int timezone)
        {
            Debug.Assert(utcDate.Kind == DateTimeKind.Utc);
            Debug.Assert(localDate.Kind == DateTimeKind.Local);

            // Check if the match time is offset by a whole number of hours (given the tolerance of TryGetHourOffset)
            if (TryGetTimezoneOffset(localDate, utcDate, out timezone)) return true;

            // If the media has a duration, compare against both the start and end.
            // Certain timestamp events correspond to one or the other and without
            // incorporating duration, the tolerance might not be satisfied.
            if (duration.HasValue)
            {
                if (TryGetTimezoneOffset(localDate, utcDate.Add(duration.Value), out timezone)) return true;
                if (TryGetTimezoneOffset(localDate, utcDate.Subtract(duration.Value), out timezone)) return true;
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
        public static PROPERTYKEY Make = new PROPERTYKEY("14b81da1-0135-4d31-96d9-6cbfc9671a99", 271); // System.Photo.CameraManufacturer
        public static PROPERTYKEY Model = new PROPERTYKEY("14b81da1-0135-4d31-96d9-6cbfc9671a99", 272); // System.Photo.CameraModel
    }

    /// <summary>
    /// A pattern to be used when renaming files so that media gets presented
    /// in the order that they were taken. Generally this is done to interleave
    /// photos and video when they use different prefixes.
    /// </summary>
    class RenamePattern
    {
        Regex m_rx;
        string m_replacement;

        /// <summary>
        /// Construct a RenamePattern
        /// </summary>
        /// <param name="regex">The Regex pattern to match.</param>
        /// <param name="replacement">The replacement pattern.</param>
        public RenamePattern(string regex, string replacement)
        {
            m_rx = new Regex(regex, RegexOptions.Singleline | RegexOptions.CultureInvariant);
            m_replacement = replacement;
        }

        /// <summary>
        /// If the pattern matches, get the new name for the file.
        /// </summary>
        /// <param name="filename">A filename to match to the pattern.</param>
        /// <param name="newName">The new name to set.</param>
        /// <returns></returns>
        public bool TryGetNewName(string filename, out string newName)
        {
            newName = m_rx.Replace(filename, m_replacement, 1);
            return !string.Equals(filename, newName, StringComparison.Ordinal);
        }

        static public bool TryGetNewName(IEnumerable<RenamePattern> list, string filename, out string newName)
        {
            foreach (var pattern in list)
            {
                if (pattern.TryGetNewName(filename, out newName)) return true;
            }
            newName = filename;
            return false;
        }
    }

}
