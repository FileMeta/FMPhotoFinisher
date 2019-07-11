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
    class PhotoFinisher
    {
        const string c_dcimDirectory = "DCIM";
        const string c_dcimTestDirectory = "DCIM_Test";

        public PhotoFinisher()
        {
            AddKeywords = new List<string>();
        }

        // Source
        SourceType m_sourceType;
        string m_sourcePath;

        // Selected Files
        List<ProcessFileInfo> m_selectedFiles = new List<ProcessFileInfo>();
        HashSet<string> m_selectedFilesHash = new HashSet<string>();
        long m_selectedFilesSize;
        int m_skippedFiles;

        // DCF Directories from which files were selected
        // This supports later cleanup of empty directories when move is used.
        List<string> m_dcfDirectories = new List<string>();

        // Deduplicate hashset
        HashSet<Guid> m_duplicateHash;
        int m_duplicatesRemoved;

        // Progress Reporting

        /// <summary>
        /// Reports messages about operations starting or completing - for the user to read.
        /// Typically shown in a log format.
        /// </summary>
        public event EventHandler<ProgressEventArgs> ProgressReported;

        void OnProgressReport(string message)
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

        void OnStatusReport(string message)
        {
            StatusReported?.Invoke(this, new ProgressEventArgs(message));
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
        /// Auto-rotate images to the vertical position.
        /// </summary>
        public bool AutoRotate { get; set; }

        /// <summary>
        /// Move files (instead of copying them) to the destination directory.
        /// </summary>
        public bool Move { get; set; }

        /// <summary>
        /// Remove duplicates or, when copying, do not copy duplcates.
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
        public bool AutoSort { get; set; }

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

        /// <summary>
        /// Select files for processing using directory path with wildcards.
        /// </summary>
        /// <param name="path">The path (including possible wildcards) of files to sleect.</param>
        /// <param name="recursive">If true, traverse subdirectories for matching files as well.</param>
        /// <remarks>
        /// <para>Wildcards may only appear in the filename, not in the earlier parts of the path.
        /// </para>
        /// </remarks>
        public void SelectFiles(string path, bool recursive)
        {
            ValidateSelectFiles(path);
            m_sourceType = recursive ? SourceType.FilePatternRecursive : SourceType.FilePattern;
            m_sourcePath = path;
        }

        /// <summary>
        /// Selects all media files from attached DCF devices.
        /// </summary>
        /// <returns>The number of files selected.</returns>
        /// <remarks>
        /// <para>Finds each removable storage device (such as a flash drive or an attached camera).
        /// For each device, looks for a "DCIM" folder according to the DCF (Design Rule for Camera
        /// File System) specification. If found, traverses all subfolders and selects the associated
        /// files.
        /// </para>
        /// </remarks>
        public void SelectDcimFiles()
        {
            m_sourceType = SourceType.DCIM;
            m_sourcePath = null;
        }

        /// <summary>
        /// Limits selected files to those that occur after the specified date/time
        /// </summary>
        public DateTime? SelectAfter { get; set; }

        /// <summary>
        /// Selects files that have changed since the last SelectIncremental on the same
        /// source and destination.
        /// </summary>
        public bool SelectIncremental { get; set; }

        /// <summary>
        /// For testing purposes. Copies files from a DCIM_Test directory into the DCIM
        /// directory.
        /// </summary>
        /// <returns>The number of files copied.</returns>
        /// <remarks>
        /// <para>This facilitates testing the <see cref="Move"/> feature by copying a set
        /// of test files into the DCIM directory before they are moved out.
        /// </para>
        /// <para>For this method to do anything, a removable drive (e.g. USB or SD) must
        /// be present with both DCIM and DCIM_Test directories on it. To be an effective
        /// test - CopyDcimTestFiles must be invoked before SelectDcfFiles.
        /// </para>
        /// </remarks>
        public int CopyDcimTestFiles()
        {
            int count = 0;

            // Process each removable drive
            foreach (DriveInfo drv in DriveInfo.GetDrives())
            {
                if (drv.IsReady && drv.DriveType == DriveType.Removable)
                {
                    // File system structure is according to JEITA "Design rule for Camera File System (DCF) which is JEITA specification CP-3461
                    // See if the DCIM folder exists
                    DirectoryInfo dcim = new DirectoryInfo(Path.Combine(drv.RootDirectory.FullName, c_dcimDirectory));
                    DirectoryInfo dcimTest = new DirectoryInfo(Path.Combine(drv.RootDirectory.FullName, c_dcimTestDirectory));
                    if (dcim.Exists && dcimTest.Exists)
                    {
                        count += CopyRecursive(dcimTest, dcim);
                    }
                }
            }

            return count;
        }

        private static int CopyRecursive(DirectoryInfo src, DirectoryInfo dst)
        {
            if (!dst.Exists) dst.Create();

            int count = 0;

            foreach (var fi in src.GetFiles())
            {
                string dstName = Path.Combine(dst.FullName, fi.Name);
                if (!File.Exists(dstName) && !Directory.Exists(dstName))
                {
                    fi.CopyTo(dstName);
                    ++count;
                }
            }

            foreach (var di in src.GetDirectories())
            {
                count += CopyRecursive(di, new DirectoryInfo(Path.Combine(dst.FullName, di.Name)));
            }

            return count;
        }

        public void PerformSelection()
        {
            switch (m_sourceType)
            {
                case SourceType.FilePattern:
                    {
                        OnProgressReport($"Selecting from: {m_sourcePath}");
                        InternalSelectFiles(m_sourcePath, false);
                        if (m_skippedFiles == 0)
                            OnProgressReport($"   {m_selectedFiles.Count} media files selected.");
                        else
                            OnProgressReport($"   {m_selectedFiles.Count} media files selected; {m_skippedFiles} skipped.");

                    }
                    break;

                case SourceType.FilePatternRecursive:
                    {
                        OnProgressReport($"Selecting tree: {m_sourcePath}");
                        InternalSelectFiles(m_sourcePath, true);
                        if (m_skippedFiles == 0)
                            OnProgressReport($"   {m_selectedFiles.Count} media files selected.");
                        else
                            OnProgressReport($"   {m_selectedFiles.Count} media files selected; {m_skippedFiles} skipped.");
                    }
                    break;

                case SourceType.DCIM:
                    {
                        InternalSelectDcimFiles();
                        OnProgressReport($"Selected {m_selectedFiles.Count} from removable camera devices.");
                    }
                    break;
            }

        }

        public void PerformOperations()
        {
            m_selectedFilesHash = null; // No longer needed once selection is complete

            // Move or copy the files if a destination directory is specified
            if (DestinationDirectory != null)
            {
                CopyOrMoveFiles();

                if (Move && m_dcfDirectories.Count > 0)
                {
                    CleanupDcfDirectories();
                }
            }

            if (DeDuplicate)
            {
                m_duplicateHash = new HashSet<Guid>();  // Required if de-duplicating
            }

            ProcessMediaFiles();

            MediaFile.DisposeOfStaticResources();

            OnProgressReport(null);
            if (m_duplicatesRemoved != 0)
            {
                OnProgressReport($"{m_duplicatesRemoved} Duplicates Removed.");
            }
            OnProgressReport($"{m_selectedFiles.Count - m_duplicatesRemoved} Files Processed.");

            OnProgressReport("All operations complete!");
        }

        void ProcessMediaFiles()
        {
            // We index through the files because we may need the index to match metadata
            // on a preceding or succeeding file
            int count = m_selectedFiles.Count;
            for (int i = 0; i < count; ++i)
            {
                var fi = m_selectedFiles[i];
                ProcessMediaFile(fi, i);
            }
        }

        void ProcessMediaFile(ProcessFileInfo fi, int index)
        {
            OnProgressReport(Path.GetFileName(fi.Filepath));

            try
            {
                // First, check for duplicate
                if (DeDuplicate)
                {
                    var hash = CalculateMd5Hash(fi.Filepath);
                    if (!m_duplicateHash.Add(hash))
                    {
                        OnProgressReport("   Removing Duplicate");
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
                        OnProgressReport("   Rename to: " + Path.GetFileName(mdf.Filepath));
                    }

                    if (SaveOriginalFilaname)
                    {
                        if (mdf.SaveOriginalFilename())
                        {
                            OnProgressReport("   Original filename saved.");
                        }
                    }

                    if (SetUuid)
                    {
                        if (mdf.SetUuid())
                        {
                            OnProgressReport("   Set UUID.");
                        }
                    }

                    bool hasCreationDate = mdf.DeterimineCreationDate(AlwaysSetDate);
                    if (hasCreationDate)
                    {
                        OnProgressReport($"   Date: {mdf.CreationDate} ({mdf.CreationDate.Date.Kind}) from {mdf.CreationDateSource}.");
                    }

                    bool hasTimezone = mdf.DetermineTimezone(AlwaysSetDate);
                    if (hasTimezone)
                    {
                        OnProgressReport($"   Timezone: {Format(mdf.Timezone)} from {mdf.TimezoneSource}.");
                    }

                    if (AutoRotate && mdf.Orientation != 1)
                    {
                        OnProgressReport("   Autorotate");
                        mdf.RotateToVertical();
                    }

                    if (Transcode && !mdf.IsPreferredFormat)
                    {
                        OnProgressReport($"   Transcode to: {mdf.PreferredFormat} ({mdf.Duration.ToString(@"hh\:mm\:ss")})");
                        if (mdf.TranscodeToPreferredFormat(msg => OnStatusReport(msg)))
                        {
                            fi.Filepath = mdf.Filepath;
                            OnProgressReport($"      Transcoded to: {Path.GetFileName(mdf.Filepath)}");
                        }
                        else
                        {
                            OnProgressReport("      Transcode failed; original format retained.");
                        }
                    }

                    if (SetDateTo != null)
                    {
                        mdf.SetDate(SetDateTo);
                        OnProgressReport($"   Date set to: {SetDateTo}");
                    }

                    if (ShiftDateBy.HasValue && mdf.ShiftDate(ShiftDateBy.Value))
                    {
                        OnProgressReport($"   Date shifted to: {mdf.CreationDate}");
                    }

                    if (SetTimezoneTo != null)
                    {
                        bool dstActive;
                        if (!hasCreationDate)
                        {
                            OnProgressReport("   ERROR: Cannot set timezone; file does not have a creationDate set.");
                        }
                        else if (mdf.SetTimezone(SetTimezoneTo, out dstActive))
                        {
                            OnProgressReport($"   Timezone set to: {mdf.Timezone} {(dstActive ? "(DST)" : "(Standard)")}");
                            // Change to forceLocal date to suppress timezone part when rendering
                            var localDate = mdf.CreationDate;
                            localDate = new FileMeta.DateTag(localDate.Date, FileMeta.TimeZoneTag.ForceLocal, localDate.Precision);
                            OnProgressReport($"   Date: {localDate} (Local)");
                        }
                        else
                        {
                            OnProgressReport($"   ERROR: Failed to set timezone.");
                        }
                    }

                    if (ChangeTimezoneTo != null)
                    {
                        bool dstActive;
                        if (!hasCreationDate)
                        {
                            OnProgressReport("   ERROR: Cannot change timezone; file does not have a creationDate set.");
                        }
                        else if (mdf.Timezone == null || mdf.Timezone.Kind != FileMeta.TimeZoneKind.Normal)
                        {
                            OnProgressReport("   ERROR: Cannot change timezone; file does not have an existing timezone. Use -setTimezone first.");
                        }
                        else if (mdf.ChangeTimezone(ChangeTimezoneTo, out dstActive))
                        {
                            OnProgressReport($"   Timezone changed to: {mdf.Timezone} {(dstActive ? "(DST)" : "(Standard)")}");
                            // Change to forceLocal date to supporess timezone part when rendering
                            var localDate = mdf.CreationDate;
                            localDate = new FileMeta.DateTag(localDate.Date, FileMeta.TimeZoneTag.ForceLocal, localDate.Precision);
                            OnProgressReport($"   Date: {localDate} (Local)");
                        }
                        else
                        {
                            OnProgressReport($"   ERROR: Failed to change timezone.");
                        }
                    }

                    if (AddKeywords.Count > 0)
                    {
                        if (mdf.AddKeywords(AddKeywords))
                        {
                            OnProgressReport($"   Tag(s) Added.");
                        }
                    }

                    if (mdf.CommitMetadata())
                    {
                        OnProgressReport("   Metadata updated.");
                    }

                    if (SetMetadataNames && mdf.SetMetadataName())
                    {
                        fi.Filepath = mdf.Filepath;
                        OnProgressReport("   Rename to: " + Path.GetFileName(mdf.Filepath));
                    }

                    if (AutoSort && !string.IsNullOrEmpty(DestinationDirectory))
                    {
                        if (hasCreationDate)
                        {
                            if (mdf.MoveFileToDatePath(DestinationDirectory))
                            {
                                fi.Filepath = mdf.Filepath;
                                OnProgressReport("   AutoSorted to: " + Path.GetDirectoryName(mdf.Filepath));
                            }
                            else
                            {
                                OnProgressReport("   Existing filename conflicts with autosort path.");
                            }
                        }
                        else
                        {
                            OnProgressReport("   No date determined - cannot autosort.");
                        }
                    }

                    if (UpdateFileSystemDateCreated && mdf.UpdateFileSystemDateCreated())
                    {
                        OnProgressReport("   Updated FS Date Created.");
                    }

                    if (UpdateFileSystemDateModified && mdf.UpdateFileSystemDateModified())
                    {
                        OnProgressReport("   Updated FS Date Modified.");
                    }

                }
            }
            catch (Exception err)
            {
                OnProgressReport($"   Error: {err.Message.Trim()}");
            }
        }

        // Throws an exception if the path isn't valid
        private void ValidateSelectFiles(string path)
        {
            string directory;
            string pattern;
            ParseSelectFilesPath(path, out directory, out pattern);
        }

        private void ParseSelectFilesPath(string path, out string directory, out string pattern)
        {
            // If has wildcards, separate the parts
            if (HasWildcard(path))
            {
                directory = Path.GetDirectoryName(path);
                pattern = Path.GetFileName(path);
                if (HasWildcard(directory))
                {
                    throw new ArgumentException($"Source '{path}' is invalid. Wildcards not allowed in directory name.");
                }
                if (string.IsNullOrEmpty(directory))
                {
                    directory = Environment.CurrentDirectory;
                }
            }

            else if (Directory.Exists(path))
            {
                directory = path;
                pattern = "*";
            }

            else
            {
                directory = Path.GetDirectoryName(path);
                pattern = Path.GetFileName(path);
                if (string.IsNullOrEmpty(directory))
                {
                    directory = Environment.CurrentDirectory;
                }
                if (!Directory.Exists(directory))
                {
                    throw new ArgumentException($"Source '{path}' does not exist.");
                }
            }
        }

        // <summary>
        // Internal selectfiles - takes count by reference so that the progress report can
        // use the total count.
        // </summary>
        private void InternalSelectFiles(string path, bool recursive)
        {
            string directory;
            string pattern;

            // Determine the "after" threshold from SelectAfter and SelectIncremental
            DateTime? after = SelectAfter;
            if (SelectIncremental)
            {
                var bookmark = new IncrementalBookmark(DestinationDirectory);
                var incrementalAfter = bookmark.GetBookmark(path);
                if (incrementalAfter.HasValue)
                {
                    if (!after.HasValue || after.Value < incrementalAfter.Value)
                    {
                        after = incrementalAfter;
                    }
                }
            }

            // Determine the path being selected
            ParseSelectFilesPath(path, out directory, out pattern);

            DateTime newestSelection = after ?? DateTime.MinValue;

            try
            {
                DirectoryInfo di = new DirectoryInfo(directory);
                foreach (var fi in di.EnumerateFiles(pattern, recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly))
                {
                    if (((m_selectedFiles.Count + m_skippedFiles) % 100) == 0)
                    {
                        string message = (m_skippedFiles == 0)
                            ? $"Selected: {m_selectedFiles.Count}"
                            : $"Selected: {m_selectedFiles.Count} Not Selected: {m_skippedFiles}";
                        OnStatusReport(message);
                    }

                    if (MediaFile.IsSupportedMediaType(fi.Extension))
                    {
                        // Limit date window of files to be selected
                        if (after.HasValue)
                        {
                            var date = MediaFile.GetBookmarkDate(fi.FullName);
                            if (!date.HasValue || date.Value <= after.Value)
                            {
                                ++m_skippedFiles;
                                continue;
                            }

                            // For this operation, we do everything in localtime because photo
                            // DateTaken metadata is in localtime.
                            // This is after-the-fact but since it's a debugging test that's OK.
                            Debug.Assert(date.Value.Kind == DateTimeKind.Local);

                            if (newestSelection < date.Value)
                                newestSelection = date.Value;
                        }

                        if (m_selectedFilesHash.Add(fi.FullName.ToLowerInvariant()))
                        {
                            m_selectedFiles.Add(new ProcessFileInfo(fi));
                            m_selectedFilesSize += fi.Length;
                        }
                    }
                }
            }
            catch (Exception err)
            {
                throw new ArgumentException($"Source '{path}' not found. ({err.Message})", err);
            }
            OnStatusReport(null);

            // If SelectIncremental, write out the new bookmark
            if (SelectIncremental && m_selectedFiles.Count > 0)
            {
                Debug.Assert(newestSelection > DateTime.MinValue);
                OnProgressReport($"   Newest: {newestSelection:yyyy'-'MM'-'dd' 'HH':'mm':'ss}");
                var bookmark = new IncrementalBookmark(DestinationDirectory);
                bookmark.SetBookmark(path, newestSelection);
            }
        }

        private void InternalSelectDcimFiles()
        {
            List<string> sourceFolders = new List<string>();

            // Process each removable drive
            foreach (DriveInfo drv in DriveInfo.GetDrives())
            {
                if (drv.IsReady && drv.DriveType == DriveType.Removable)
                {
                    try
                    {
                        // File system structure is according to JEITA "Design rule for Camera File System (DCF) which is JEITA specification CP-3461
                        // See if the DCIM folder exists
                        DirectoryInfo dcim = new DirectoryInfo(Path.Combine(drv.RootDirectory.FullName, c_dcimDirectory));
                        if (dcim.Exists)
                        {
                            // Folders containing images must be named with three digits followed
                            // by five alphanumeric characters. First digit cannot be zero.
                            foreach (DirectoryInfo di in dcim.EnumerateDirectories())
                            {
                                if (di.Name.Length == 8)
                                {
                                    int dirnum;
                                    if (int.TryParse(di.Name.Substring(0, 3), out dirnum) && dirnum >= 100 && dirnum <= 999)
                                    {
                                        sourceFolders.Add(di.FullName);
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception)
                    {
                        // Suppress the error and move to the next drive.
                    }
                } // If drive is ready and removable
            } // for each drive

            foreach (var path in sourceFolders)
            {
                int bookmark = m_selectedFiles.Count;
                InternalSelectFiles(path, false);
                if (m_selectedFiles.Count > bookmark) m_dcfDirectories.Add(path);
            }
        }

        private void CopyOrMoveFiles()
        {
            string verb = Move ? "Moving" : "Copying";
            string dstDirectory = DestinationDirectory;
            // TODO: If -sort then add a temporary directory.

            OnProgressReport($"{verb} media files to working folder: {dstDirectory}.");

            uint startTicks = (uint)Environment.TickCount;
            long bytesCopied = 0;

            int n = 0;
            foreach (var fi in m_selectedFiles)
            {
                if (bytesCopied == 0)
                {
                    OnStatusReport($"{verb} file {n + 1} of {m_selectedFiles.Count}");
                }
                else
                {
                    uint ticksElapsed;
                    unchecked
                    {
                        ticksElapsed = (uint)Environment.TickCount - startTicks;
                    }

                    double bps = ((double)bytesCopied * 1000.0) / (double)ticksElapsed;
                    double remaining = (m_selectedFilesSize - bytesCopied) / bps;
                    TimeSpan remain = new TimeSpan(((long)((m_selectedFilesSize - bytesCopied) / bps)) * 10000000L);

                    OnStatusReport($"{verb} file {n + 1} of {m_selectedFiles.Count}. Time remaining: {remain.FmtCustom()} MBps: {(bps / (1024 * 1024)):#,###.###}");
                }

                string dstFilepath = Path.Combine(DestinationDirectory, Path.GetFileName(fi.OriginalFilepath));
                MediaFile.MakeFilepathUnique(ref dstFilepath);

                if (Move)
                {
                    File.Move(fi.Filepath, dstFilepath);
                }
                else
                {
                    File.Copy(fi.Filepath, dstFilepath);
                }
                fi.Filepath = dstFilepath;
                bytesCopied += fi.Size;

                ++n;
            }

            TimeSpan elapsed;
            unchecked
            {
                uint ticksElapsed = (uint)Environment.TickCount - startTicks;
                elapsed = new TimeSpan(ticksElapsed * 10000L);
            }

            OnStatusReport(null);
            OnProgressReport($"{verb} complete. {m_selectedFiles.Count} files, {bytesCopied / (1024.0 * 1024.0): #,##0.0} MB, {elapsed.FmtCustom()} elapsed");
        }

        /// <summary>
        /// Remove empty DCF directories from which files were moved.
        /// </summary>
        private void CleanupDcfDirectories()
        {
            foreach (string directoryName in m_dcfDirectories)
            {
                // Clean up the folders
                try
                {
                    DirectoryInfo di = new DirectoryInfo(directoryName);

                    // Get rid of thumbnails unless a matching file still exists
                    foreach (FileInfo fi in di.GetFiles("*.thm"))
                    {
                        bool hasMatch = false;
                        foreach (FileInfo fi2 in di.GetFiles(Path.GetFileNameWithoutExtension(fi.Name) + ".*"))
                        {
                            if (!string.Equals(fi2.Extension, ".thm", StringComparison.OrdinalIgnoreCase))
                            {
                                hasMatch = true;
                                break;
                            }
                        }
                        if (!hasMatch) fi.Delete();
                    }

                    // Get rid of Windows thumbnails file (if it exists)
                    {
                        string thumbName = Path.Combine(di.FullName, "Thumbs.db");
                        if (File.Exists(thumbName)) File.Delete(thumbName);
                    }

                    // If the folder is empty, delete it
                    if (!di.EnumerateFileSystemInfos().Any()) di.Delete();
                }
                catch (Exception err)
                {
                    // Report errors during cleanup but proceed with other files.
                    OnProgressReport($"Error cleaning up folders on removable drive '{Path.GetPathRoot(directoryName)}': {err.Message}");
                }
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

        private static char[] s_wildcards = new char[] { '*', '?' };

        private static bool HasWildcard(string filename)
        {
            return filename.IndexOfAny(s_wildcards) >= 0;
        }

        private static string Format(FileMeta.TimeZoneTag tz)
        {
            if (tz.Kind != FileMeta.TimeZoneKind.Normal) return $"({tz.Kind})";
            return tz.ToString();
        }

        private enum SourceType
        {
            FilePattern = 1,
            FilePatternRecursive = 2,
            DCIM = 3
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
