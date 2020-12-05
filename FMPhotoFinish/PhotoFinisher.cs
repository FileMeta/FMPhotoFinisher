using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Diagnostics;
using System.Globalization;
using System.Text.RegularExpressions;

namespace FMPhotoFinish
{
    enum SetMode : int
    {
        DoNothing = 0,  // Don't write any value
        SetIfEmpty = 1, // Set if the value is empty
        SetAlways = 2   // Overwrite any existing value
    }
    class PhotoFinisher : IMediaQueue
    {
        public PhotoFinisher()
        {
            AddKeywords = new List<string>();
        }

        // Selected Files
        List<ProcessFileInfo> m_fileQueue = new List<ProcessFileInfo>();

        // Deduplicate hashset
        HashSet<Guid> m_duplicateHash;
        int m_duplicatesRemoved;

        // Progress Reporting

        /// <summary>
        /// Reports messages about operations starting or completing - for the user to read.
        /// Typically shown in a log format.
        /// </summary>
        public event EventHandler<ProgressEventArgs> ProgressReported;

        public void ReportProgress(string message)
        {
            ProgressReported?.Invoke(this, new ProgressEventArgs(message));
        }

        /// <summary>
        /// Reports messages about partial completion (e.g. "3 of 5 copied").
        /// </summary>
        /// <remarks>
        /// This would typically be shown in a status bar. When sent to the console, it should be sent
        /// to stderr with a \r (not \r\n). A null or blank message should clear the status.
        /// </remarks>
        public event EventHandler<ProgressEventArgs> StatusReported;

        public void ReportStatus(string message)
        {
            StatusReported?.Invoke(this, new ProgressEventArgs(message));
        }

        public void Add(ProcessFileInfo pfi)
        {
            m_fileQueue.Add(pfi);
        }

        #region Operations

        /// <summary>
        /// For canon cameras, changes name prefixes to all be IMG - thereby ensuring they sort in order.
        /// </summary>
        public bool SetOrderedNames { get; set; }

        /// <summary>
        /// Change the filename to one based on the date the photo was taken plus subject and title metadata.
        /// </summary>
        public bool SetMetadataNames { get; set; }

        /// <summary>
        /// Update subject, title, and metadata keywords from the filename
        /// </summary>
        public SetMode MetadataFromFilename { get; set; }

        /// <summary>
        /// Auto-rotate images to the vertical position.
        /// </summary>
        public bool AutoRotate { get; set; }

        /// <summary>
        /// Remove duplicates or, when copying, do not copy duplicates.
        /// </summary>
        public bool DeDuplicate { get; set; }

        /// <summary>
        /// Set the Date to the specified value
        /// </summary>
        public FileMeta.DateTag SetDateTo { get; set; }

        /// <summary>
        /// Shift the Date by the specified timespan.
        /// </summary>
        public TimeSpan? ShiftDateBy { get; set; }

        /// <summary>
        /// Determine the timezone and set the custom metadata field accordingly
        /// </summary>
        /// <remarks>
        /// If other metadata are written then the timezone will also be written. This
        /// forces the field to be written even if nothing else changes.
        /// </remarks>
        public bool AlwaysSetDate { get; set; }

        /// <summary>
        /// Sets the timezone to the specified value while keeping the local time the same.
        /// </summary>
        /// <remarks>
        /// May be combined with ChangeTimezoneTo in which case set timezone will occur first
        /// and then change timezone.
        /// </remarks>
        public TimeZoneInfo SetTimezoneTo { get; set; }

        /// <summary>
        /// Changes the timezone to the specified value while keeping the UTC time the same.
        /// </summary>
        /// <remarks>
        /// May be combined with SetTimezoneTo in which case set timezone will occur first
        /// and then change timezone.
        /// </remarks>
        public TimeZoneInfo ChangeTimezoneTo { get; set; }

        /// <summary>
        /// Updates the file system dateCreated value to match the metadata date created.
        /// </summary>
        public bool UpdateFileSystemDateCreated { get; set; }

        /// <summary>
        /// Updates the file system dateModified value to match the metadata date created.
        /// </summary>
        public bool UpdateFileSystemDateModified { get; set; }

        /// <summary>
        /// Transcode audio and video files into the preferred format. Also renames .jpeg to .jpg
        /// </summary>
        public bool Transcode { get; set; }

        /// <summary>
        /// Auto-sort the images into directories according to date taken
        /// </summary>
        public DatePathType SortBy { get; set; }

        /// <summary>
        /// Save the original filename in a custom metaTag (if it hasn't already been stored).
        /// </summary>
        public bool SaveOriginalFilaname { get; set; }

        /// <summary>
        /// Set a UUID in a custom metatag (if one doesn't already exist)
        /// </summary>
        public bool SetUuid { get; set; }

        /// <summary>
        /// Add keyword tags to the file
        /// </summary>
        public IList<string> AddKeywords { get; private set; }

        /// <summary>
        /// The destination directory - files will be copied or moved there.
        /// </summary>
        public string DestinationDirectory { get; set; }

        #endregion Operations

        public void PerformOperations()
        {
            if (DeDuplicate)
            {
                m_duplicateHash = new HashSet<Guid>();  // Required if de-duplicating
            }

            ProcessMediaFiles();

            MediaFile.DisposeOfStaticResources();

            ReportProgress(null);
            if (m_duplicatesRemoved != 0)
            {
                ReportProgress($"{m_duplicatesRemoved} Duplicates Removed.");
            }
            ReportProgress($"{m_fileQueue.Count - m_duplicatesRemoved} Files Processed.");

            ReportProgress("All operations complete!");
        }

        void ProcessMediaFiles()
        {
            foreach(var fi in m_fileQueue)
            {
                ProcessMediaFile(fi);
            }
        }

        void ProcessMediaFile(ProcessFileInfo fi)
        {
            ReportProgress(Path.GetFileName(fi.Filepath));

            try
            {
                // First, check for duplicate
                if (DeDuplicate)
                {
                    var hash = CalculateMd5Hash(fi.Filepath);
                    if (!m_duplicateHash.Add(hash))
                    {
                        ReportProgress("   Removing Duplicate");
                        File.Delete(fi.Filepath);
                        ++m_duplicatesRemoved;
                        return;
                    }
                }

                using (var mdf = new MediaFile(fi.Filepath, Path.GetFileName(fi.OriginalFilepath)))
                {
                    mdf.OriginalDateCreated = fi.OriginalDateCreated;
                    mdf.OriginalDateModified = fi.OriginalDateModified;

                    if (SetOrderedNames && mdf.SetOrderedName())
                    {
                        fi.Filepath = mdf.Filepath;
                        ReportProgress("   Rename to: " + Path.GetFileName(mdf.Filepath));
                    }

                    if (SaveOriginalFilaname)
                    {
                        if (mdf.SaveOriginalFilename())
                        {
                            ReportProgress("   Original filename saved.");
                        }
                    }

                    if (SetUuid)
                    {
                        if (mdf.SetUuid())
                        {
                            ReportProgress("   Set UUID.");
                        }
                    }

                    bool hasCreationDate = mdf.DeterimineCreationDate(AlwaysSetDate);
                    if (hasCreationDate)
                    {
                        ReportProgress($"   Date: {mdf.CreationDate} ({mdf.CreationDate.Date.Kind}) from {mdf.CreationDateSource}.");
                    }

                    bool hasTimezone = mdf.DetermineTimezone(AlwaysSetDate);
                    if (hasTimezone)
                    {
                        ReportProgress($"   Timezone: {Format(mdf.Timezone)} from {mdf.TimezoneSource}.");
                    }

                    if (Transcode && !mdf.IsPreferredFormat)
                    {
                        ReportProgress($"   Transcode to: {mdf.PreferredFormat} ({mdf.Duration.ToString(@"hh\:mm\:ss")})");
                        if (mdf.TranscodeToPreferredFormat(msg => ReportStatus(msg)))
                        {
                            fi.Filepath = mdf.Filepath;
                            ReportProgress($"      Transcoded to: {Path.GetFileName(mdf.Filepath)}");
                        }
                        else
                        {
                            ReportProgress("      Transcode failed; original format retained.");
                        }
                    }

                    if (AutoRotate && mdf.Orientation != 1)
                    {
                        ReportProgress("   Autorotate");
                        mdf.RotateToVertical();
                    }

                    if (SetDateTo != null)
                    {
                        mdf.SetDate(SetDateTo);
                        ReportProgress($"   Date set to: {SetDateTo}");
                    }

                    if (ShiftDateBy.HasValue && mdf.ShiftDate(ShiftDateBy.Value))
                    {
                        ReportProgress($"   Date shifted to: {mdf.CreationDate}");
                    }

                    if (SetTimezoneTo != null)
                    {
                        bool dstActive;
                        if (!hasCreationDate)
                        {
                            ReportProgress("   ERROR: Cannot set timezone; file does not have a creationDate set.");
                        }
                        else if (mdf.SetTimezone(SetTimezoneTo, out dstActive))
                        {
                            ReportProgress($"   Timezone set to: {mdf.Timezone} {(dstActive ? "(DST)" : "(Standard)")}");
                            // Change to forceLocal date to suppress timezone part when rendering
                            var localDate = mdf.CreationDate;
                            localDate = new FileMeta.DateTag(localDate.Date, FileMeta.TimeZoneTag.ForceLocal, localDate.Precision);
                            ReportProgress($"   Date: {localDate} (Local)");
                        }
                        else
                        {
                            ReportProgress($"   ERROR: Failed to set timezone.");
                        }
                    }

                    if (ChangeTimezoneTo != null)
                    {
                        bool dstActive;
                        if (!hasCreationDate)
                        {
                            ReportProgress("   ERROR: Cannot change timezone; file does not have a creationDate set.");
                        }
                        else if (mdf.Timezone == null || mdf.Timezone.Kind != FileMeta.TimeZoneKind.Normal)
                        {
                            ReportProgress("   ERROR: Cannot change timezone; file does not have an existing timezone. Use -setTimezone first.");
                        }
                        else if (mdf.ChangeTimezone(ChangeTimezoneTo, out dstActive))
                        {
                            ReportProgress($"   Timezone changed to: {mdf.Timezone} {(dstActive ? "(DST)" : "(Standard)")}");
                            // Change to forceLocal date to supporess timezone part when rendering
                            var localDate = mdf.CreationDate;
                            localDate = new FileMeta.DateTag(localDate.Date, FileMeta.TimeZoneTag.ForceLocal, localDate.Precision);
                            ReportProgress($"   Date: {localDate} (Local)");
                        }
                        else
                        {
                            ReportProgress($"   ERROR: Failed to change timezone.");
                        }
                    }

                    if (MetadataFromFilename != SetMode.DoNothing)
                    {
                        mdf.MetadataFromFilename(MetadataFromFilename == SetMode.SetAlways);
                    }

                    if (AddKeywords.Count > 0)
                    {
                        if (mdf.AddKeywords(AddKeywords))
                        {
                            ReportProgress($"   Tag(s) Added.");
                        }
                    }

                    if (mdf.CommitMetadata())
                    {
                        ReportProgress("   Metadata updated.");
                    }

                    if (SetMetadataNames && mdf.MetadataToFilename())
                    {
                        fi.Filepath = mdf.Filepath;
                        ReportProgress("   Rename to: " + Path.GetFileName(mdf.Filepath));
                    }

                    if (SortBy != DatePathType.None && !string.IsNullOrEmpty(DestinationDirectory))
                    {
                        if (hasCreationDate)
                        {
                            if (mdf.MoveFileToDatePath(DestinationDirectory, SortBy))
                            {
                                fi.Filepath = mdf.Filepath;
                                ReportProgress("   AutoSorted to: " + Path.GetDirectoryName(mdf.Filepath));
                            }
                            else
                            {
                                ReportProgress("   Existing filename conflicts with autosort path.");
                            }
                        }
                        else
                        {
                            ReportProgress("   No date determined - cannot autosort.");
                        }
                    }

                    if (UpdateFileSystemDateCreated && mdf.UpdateFileSystemDateCreated())
                    {
                        ReportProgress("   Updated FS Date Created.");
                    }

                    if (UpdateFileSystemDateModified && mdf.UpdateFileSystemDateModified())
                    {
                        ReportProgress("   Updated FS Date Modified.");
                    }

                }
            }
            catch (Exception err)
            {
                ReportProgress($"   Error: {err.Message.Trim()}");
            }
        }

        private static Guid CalculateMd5Hash(string filepath)
        {
            byte[] hash;
            using (var stream = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                MD5 hasher = MD5.Create();
                hash = hasher.ComputeHash(stream);
            }
            Debug.Assert(hash.Length == 16);
            return new Guid(hash);  // Conveniently, a GUID is the same size as an MD5 hash - 128 bits
        }

        private static string Format(FileMeta.TimeZoneTag tz)
        {
            if (tz.Kind != FileMeta.TimeZoneKind.Normal) return $"({tz.Kind})";
            return tz.ToString();
        }

    } // Class PhotoFinisher

    class ProgressEventArgs : EventArgs
    {
        public ProgressEventArgs(string message)
        {

            Message = message;
        }

        public string Message { get; private set; }
    }

} // Namespace
