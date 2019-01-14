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
        const string c_datePrecisionKey = "datePrecision";
        const string c_originalFilenameKey = "originalFilename";
        const string c_uuidKey = "uuid";
        //const string c_makeKey = "make";
        //const string c_modelKey = "model";

        const string c_defaultTitle = "Pic";

        static readonly TimeSpan s_timespanZero = new TimeSpan(0);

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

        // File System Values (In local time)
        string m_fsOriginalFilename;      // Filename when process started.
        DateTime m_fsDateCreated;
        DateTime m_fsDateModified;

        // Values from the Windows Property System
        DateTime? m_psDateTaken;
        DateTime? m_psDateEncoded;
        TimeSpan? m_psDuration;
        string m_psSubject;
        string m_psTitle;

        // Metatag values from Property System Keywords
        TimeZoneTag m_mtTimezone;
        int m_mtDatePrecision; // Zero means no value

        // Values from IsomCoreMetadata
        DateTime? m_isomCreationTime;

        // Values from ExifTool
        DateTime? m_etDateTimeOriginal; // EXIF:DateTimeOriginal (.jpg, .jpeg) - RIFF:DateTimeOriginal (.avi) - in local time.
        TimeZoneTag m_etTimezone;

        // Master Values (to be preserved and set)
        bool m_updateMetadata = false; // Metadata should be updated upon disposal
        DateTime? m_creationDate;
        TimeZoneTag m_timezone;
        int m_datePrecision;
        string m_make;
        string m_model;
        string m_originalFilename;
        Guid m_uuid;

        public MediaFile(string filepath, string originalFilename)
        {
            m_filepath = filepath;
            m_fsOriginalFilename = originalFilename ?? Path.GetFileName(filepath);

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
                m_make = propstore.GetValue(PropertyKeys.Make) as string;
                m_model = propstore.GetValue(PropertyKeys.Model) as string;
                m_psSubject = propstore.GetValue(PropertyKeys.Subject) as string;
                m_psTitle = propstore.GetValue(PropertyKeys.Title) as string;

                if (m_psSubject != null)
                    m_psSubject = m_psSubject.Trim();

                if (m_psTitle != null)
                    m_psTitle = m_psTitle.Trim();

                // Windows property system will fill in the subject with the title if subject is not present. Compensate for that.
                if (string.Equals(m_psSubject, m_psTitle, StringComparison.Ordinal))
                    m_psSubject = null; 

                // Keywords may be used to store custom metadata
                var metaTagSet = new MetaTagSet();
                metaTagSet.LoadMetatags((string)propstore.GetValue(PropertyKeys.Comment));
                {
                    string value;

                    // Timezone
                    TimeZoneTag tz;
                    if (metaTagSet.MetaTags.TryGetValue(c_timezoneKey, out value)
                        && TimeZoneTag.TryParse(value, out tz))
                    {
                        m_mtTimezone = tz;
                    }

                    // Date precision
                    int precision;
                    if (metaTagSet.MetaTags.TryGetValue(c_datePrecisionKey, out value)
                        && int.TryParse(value, out precision)
                        && precision >= DateTag.PrecisionMin && precision <= DateTag.PrecisionMax)
                    {
                        m_mtDatePrecision = precision;
                    }

                    // Original Filename
                    if (metaTagSet.MetaTags.TryGetValue(c_originalFilenameKey, out value)
                        && !string.IsNullOrEmpty(value))
                    {
                        m_originalFilename = value;
                    }

                    // UUID
                    Guid uuid;
                    if (metaTagSet.MetaTags.TryGetValue(c_uuidKey, out value)
                        && Guid.TryParse(value, out uuid))
                    {
                        m_uuid = uuid;
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
                // Init ExifTool if it hasn't been done so far.
                if (s_exifTool == null)
                {
                    s_exifTool = new ExifToolWrapper.ExifTool();
                }

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
        public TimeSpan Duration { get { return m_psDuration ?? s_timespanZero; } }

        public DateTag CreationDate
        {
            get
            {
                DateTime date = m_creationDate ?? m_psDateTaken ?? m_psDateEncoded ?? m_isomCreationTime ?? m_etDateTimeOriginal ?? m_fsDateCreated;
                return new DateTag(date, m_timezone, m_datePrecision);   // If timezone has not yet been determined, the constructor fills in a default value.
            }
        }

        public int DatePrecision
        {
            get { return (m_datePrecision > 0) ? m_datePrecision : m_mtDatePrecision; }
        }

        public string CreationDateSource { get; private set; }

        public TimeZoneTag Timezone
        {
            get { return m_timezone ?? m_mtTimezone ?? m_etTimezone ?? null; }
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
            if (RenamePattern.TryGetNewName(s_renamePatterns, m_fsOriginalFilename, out newName))
            {
                // Create the new path and make it unique
                string newPath = Path.Combine(Path.GetDirectoryName(m_filepath), newName);
                MakeFilepathUnique(ref newPath);

                // Rename
                File.Move(m_filepath, newPath);
                m_filepath = newPath;
                return true;
            }

            return false;
        }

        /// <summary>
        /// Change the filename to be based on the date the photo was taken plus subject and title metadata
        /// </summary>
        /// <returns>True if the name was changed. False if the photo has no DateTaken metadata.</returns>
        /// <remarks>
        /// <para>The new filename pattern is: yyyy-mm-dd_hhmmss &lt;subject&gt; - &lt;title&gt;.jpg.
        /// For example, "2019-01-15_142022 Mirror Lake - John Fishing.jpg".
        /// </para>
        /// <para>If the title is not present, no dash or title will appear. If the subject is not
        /// present, a dash and the title will appear. If neither is present, the word, "Pic" will
        /// substitute.
        /// </para>
        /// <para>If more than one photo was taken at the exact same second then the filename will have a
        /// numeric extension like this: "2019-01-15_142022_Pic (01).jpg.
        /// </para>
        /// </remarks>
        public bool SetMetadataName()
        {
            if (!m_creationDate.HasValue) return false;

            // Get a local dateTime for the file
            var dt = CreationDate.ResolveTimeZone(TimeZoneInfo.Local).ToLocal();

            string newName = dt.ToString("yyyy-MM-dd_HHmmss",
                System.Globalization.CultureInfo.InvariantCulture);

            if (!string.IsNullOrEmpty(m_psSubject))
                newName = string.Concat(newName, " ", m_psSubject);

            if (!string.IsNullOrEmpty(m_psTitle))
                newName = string.Concat(newName, " - ", m_psTitle);

            if (string.IsNullOrEmpty(m_psTitle) && string.IsNullOrEmpty(m_psSubject))
            {
                newName = string.Concat(newName, " ", c_defaultTitle);
            }

            newName = string.Concat(newName, Path.GetExtension(m_filepath));

            // Change the filename
            string dstPath = Path.Combine(Path.GetDirectoryName(m_filepath), newName);
            MakeFilepathUnique(ref dstPath);
            File.Move(m_filepath, dstPath);
            m_filepath = dstPath;

            return true;
        }

        /// <summary>
        /// Moves the file to a path based on its creation date.
        /// </summary>
        /// <returns>True if successful. False if a date has not been set or if there is a file by the name of one of the directories in the path.</returns>
        public bool MoveFileToDatePath(string dstRoot)
        {
            if (!m_creationDate.HasValue) return false;

            // Get a local dateTime for the file
            var dt = CreationDate.ResolveTimeZone(TimeZoneInfo.Local).ToLocal();

            // Create the directory string year\month\day
            // This part deliberately uses cultural-sensitive encoding so a system configured for French will use French month and day names.
            // It also uses Path.Combine for forward and backslash compatibility with non-Windows systems (though there are other dependencies in this app)
            string dir = Path.Combine(dstRoot, dt.ToString("yyyy"));
            dir = Path.Combine(dir, dt.ToString("MM MMMM"));
            dir = Path.Combine(dir, dt.ToString("dd~ddd"));

            // Create the directory (this will handle the whole path if necessary)
            // Succeeds with no error if the directory already exists
            try
            {
                Directory.CreateDirectory(dir);
            }
            catch
            {
                return false;
            }

            // Move the file to the directory
            string dstPath = Path.Combine(dir, Path.GetFileName(m_filepath));
            MakeFilepathUnique(ref dstPath);
            File.Move(m_filepath, dstPath);
            m_filepath = dstPath;

            return true;
        }

        /// <summary>
        /// Attempt to determine the date of media files - that is, the date that the event
        /// was recorded. This is the DateTimeOriginal from JPEG EXIF and the creation_date
        /// from .MOV and .MP4 files.
        /// </summary>
        /// <returns>True if the date was successfully determined. Otherwise false.</returns>
        public bool DeterimineCreationDate()
        {
            // First, determine the precision if not already set
            if (m_datePrecision == 0 && m_mtDatePrecision != 0)
            {
                m_datePrecision = m_mtDatePrecision;
            }

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
                if (TryParseDateTimeFromFilename(m_fsOriginalFilename, out dt))
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

            if (m_mtTimezone != null)
            {
                m_timezone = m_mtTimezone;
                TimezoneSource = "MetaTag";
                // Not necessary to update metadata because this is the primary source
                return true;
            }

            if (m_etTimezone != null)
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
                && TryParseDateTimeFromFilename(m_fsOriginalFilename, out dt)
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
                return true;
            }

            if (m_mediaType != MediaType.Audio && m_mediaType != MediaType.Video)
                return false;

            return Transcode(reporter);
        }

        public void SetDate(DateTag date)
        {
            m_creationDate = date.Date;
            m_timezone = date.TimeZone;
            m_datePrecision = date.Precision;
            m_updateMetadata = true;
            CreationDateSource = "Set";
            TimezoneSource = "Set";
        }

        public bool ShiftDate(TimeSpan timespan)
        {
            if (!m_creationDate.HasValue) return false;

            m_creationDate += timespan;
            m_updateMetadata = true;
            CreationDateSource = "Shifted";
            return true;
        }

        public bool SetTimezone(TimeZoneInfo tzi, out bool dstActive)
        {
            // If no creationDate, do nothing.
            if (!m_creationDate.HasValue)
            {
                dstActive = false;
                return false;
            }

            // If no timezone has yet been determined, set to ForceLocal
            if (m_timezone == null) m_timezone = TimeZoneTag.ForceLocal;

            // Using the existing timezone, make sure the creationDate is in local time (does nothing if already local)
            m_creationDate = m_timezone.ToLocal(m_creationDate.Value);

            // Set the timezone offset for the creationDate (may be different values according to daylight savings)
            dstActive = tzi.IsDaylightSavingTime(m_creationDate.Value);
            m_timezone = new TimeZoneTag(tzi.GetUtcOffset(m_creationDate.Value), TimeZoneKind.Normal);

            // When metadata is committed, m_creationDate will be stored in either local or UTC according to the file type.
            m_updateMetadata = true;
            TimezoneSource = "Set";

            return true;
        }

        public bool ChangeTimezone(TimeZoneInfo tzi, out bool dstActive)
        {
            // If no creationDate, do nothing.
            dstActive = false;
            if (!m_creationDate.HasValue) return false;

            // If no existing timezone, do nothing.
            if (m_timezone == null || m_timezone.Kind != TimeZoneKind.Normal) return false;

            // Using the existing timezone, make sure the creationDate is in UTC (does nothing if already UTC)          
            m_creationDate = m_timezone.ToUtc(m_creationDate.Value);

            // Set the timezone offset for the creationDate (may be different values according to daylight savings)
            dstActive = tzi.IsDaylightSavingTime(m_creationDate.Value);
            m_timezone = new TimeZoneTag(tzi.GetUtcOffset(m_creationDate.Value), TimeZoneKind.Normal);

            // When metadata is committed, m_creationDate will be stored in either local or UTC according to the file type.
            m_updateMetadata = true;
            TimezoneSource = "Set";

            return true;
        }

        /// <summary>
        /// Update the FileSystem dateCreated value to match the metadata.
        /// </summary>
        /// <returns>True if the FileSystem dateCreated was updated. False if the date created
        /// has not been determined and so no update was accomplished.</returns>
        /// <remarks>
        /// <para></para>
        /// </remarks>
        public bool UpdateFileSystemDateCreated()
        {
            // If no creationDate, do nothing.
            if (!m_creationDate.HasValue) return false;

            // Convert to UTC
            DateTime dateUtc = CreationDate.ResolveTimeZone(TimeZoneInfo.Local).ToUtc();

            File.SetCreationTimeUtc(m_filepath, dateUtc);
            return true;
        }

        /// <summary>
        /// Update the FileSystem dateCreated value to match the metadata.
        /// </summary>
        /// <returns>True if the FileSystem dateModified was updated. False if the date created
        /// has not been determined and so no update was accomplished.</returns>
        /// <remarks>
        /// <para></para>
        /// </remarks>
        public bool UpdateFileSystemDateModified()
        {
            // If no creationDate, do nothing.
            if (!m_creationDate.HasValue) return false;

            // Convert to UTC
            DateTime dateUtc = CreationDate.ResolveTimeZone(TimeZoneInfo.Local).ToUtc();

            File.SetLastWriteTimeUtc(m_filepath, dateUtc);
            return true;
        }

        /// <summary>
        /// Save the original filename in custom metadata field.
        /// </summary>
        /// <returns>True if original filename saved. False if original filename
        /// is already stored.</returns>
        public bool SaveOriginalFilename()
        {
            if (!string.IsNullOrEmpty(m_originalFilename)) return false;

            Debug.Assert(!string.IsNullOrEmpty(m_fsOriginalFilename));
            m_originalFilename = m_fsOriginalFilename;
            m_updateMetadata = true;
            return true;
        }

        /// <summary>
        /// Set a uuid in custom metadata field.
        /// </summary>
        /// <returns>True if stored a new UUID. False if UUID already exists.</returns>
        public bool SetUuid()
        {
            if (!Guid.Empty.Equals(m_uuid)) return false;

            m_uuid = Guid.NewGuid();
            m_updateMetadata = true;
            return true;
        }

        /// <summary>
        /// Save any metadata fields that have changed.
        /// </summary>
        /// <returns>True if metadata updated. False if no metadata changes to save.</returns>
        public bool CommitMetadata()
        {
            if (!m_updateMetadata) return false; // Nothing to update

            if (m_mediaType == MediaType.Unsupported)
                throw new ApplicationException("Cannot update metadata on unsupported media type.");

            // Timezone to use for conversions as values are saved.
            var timezone = m_timezone ?? TimeZoneTag.Zero;

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
                        var dt = timezone.ToUtc(m_creationDate.Value);
                        isom.CreationTime = dt;
                        isom.ModificationTime = dt;
                        isom.Commit();
                    }
                    creationDateStoredByIsom = true;
                }
            }

            try
            {
                using (var ps = PropertyStore.Open(m_filepath, true))
                {
                    // Prep the metatags with existing values
                    var metaTagSet = new MetaTagSet();

                    // Handle type-specific metadata
                    if (m_mediaType == MediaType.Image)
                    {

                        if (m_creationDate.HasValue && !creationDateStoredByIsom)
                        {
                            // Convert to local (this does nothing if it is already Local.
                            var dt = timezone.ToLocal(m_creationDate.Value);
                            ps.SetValue(PropertyKeys.DateTaken, dt);
                        }
                    }

                    // Audio and video both use Isom file format (.mp4 and .m4a)
                    else
                    {
                        if (m_creationDate.HasValue && !creationDateStoredByIsom)
                        {
                            // Convert to UTC (this does nothing if it is already UTC.
                            var dt = m_timezone.ToUtc(m_creationDate.Value);
                            ps.SetValue(PropertyKeys.DateEncoded, dt);
                        }
                    }

                    if (m_timezone != null)
                        metaTagSet.MetaTags[c_timezoneKey] = m_timezone.ToString();
                    if (m_datePrecision >= DateTag.PrecisionMin)
                        metaTagSet.MetaTags[c_datePrecisionKey] = m_datePrecision.ToString();
                    if (!string.IsNullOrEmpty(m_originalFilename))
                        metaTagSet.MetaTags[c_originalFilenameKey] = m_originalFilename;
                    if (!m_uuid.Equals(Guid.Empty))
                        metaTagSet.MetaTags[c_uuidKey] = m_uuid.ToString("D");
                    if (!string.IsNullOrEmpty(m_make))
                        ps.SetValue(PropertyKeys.Make, m_make);
                    if (!string.IsNullOrEmpty(m_model))
                        ps.SetValue(PropertyKeys.Model, m_model);

                    if (metaTagSet.MetaTags.Count > 0)
                    {
                        ps.SetValue(PropertyKeys.Comment,
                            metaTagSet.AddMetatagsToString(
                                ps.GetValue(PropertyKeys.Comment) as string));
                    }

                    ps.Commit();
                }
            }
            catch (Exception err)
            {
                // Translate error message
                throw new ApplicationException("Error storing metadata - file may be corrupt.", err);
            }

            m_updateMetadata = false;
            return true;
        }

        #region IDisposable Support

        protected virtual void Dispose(bool disposing)
        {
            if (m_filepath == null) return;

            try
            {
                if (disposing)
                {
                }

                else
                {
                    System.Diagnostics.Debug.Fail("Failed to dispose of MediaFile.");
                }
            }
            finally
            {
                m_filepath = null;
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
                m_updateMetadata = true; // Need to restore metadata to the transcoded file
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

#region Init and Shutdown

        static ExifToolWrapper.ExifTool s_exifTool;
        static readonly StaticDisposer s_psDisposer = new StaticDisposer();

        public static void DisposeOfStaticResources()
        {
            if (s_exifTool != null)
            {
                var exifTool = s_exifTool;
                s_exifTool = null;
                exifTool.Dispose();
            }
        }

        private sealed class StaticDisposer
        {
            ~StaticDisposer()
            {
                DisposeOfStaticResources();
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
        public static PROPERTYKEY ImageId = new PROPERTYKEY("10DABE05-32AA-4C29-BF1A-63E2D220587F", 100); // System.Image.ImageID
        public static PROPERTYKEY Comment = new PROPERTYKEY("F29F85E0-4FF9-1068-AB91-08002B27B3D9", 6); // 
        public static PROPERTYKEY Subject = new PROPERTYKEY("F29F85E0-4FF9-1068-AB91-08002B27B3D9", 3); // System.Subject
        public static PROPERTYKEY Title = new PROPERTYKEY("F29F85E0-4FF9-1068-AB91-08002B27B3D9", 2); // System.Title
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
