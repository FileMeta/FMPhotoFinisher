using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;
using System.Text.RegularExpressions;

namespace FMPhotoFinish
{
    class PhotoFinisher
    {
        // Selected Files
        List<ProcessFileInfo> m_selectedFiles = new List<ProcessFileInfo>();
        HashSet<string> m_selectedFilesHash = new HashSet<string>();
        long m_selectedFilesSize;

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
        /// Auto-rotate images to the vertical position.
        /// </summary>
        public bool AutoRotate { get; set; }

        /// <summary>
        /// Move files (instead of copying them) to the destination directory.
        /// </summary>
        public bool Move { get; set; }

        /// <summary>
        /// Auto-sort the images into directories according to date taken
        /// </summary>
        public bool AutoSort { get; set; }

        /// <summary>
        /// Transcode audio and video files into the preferred format. Also renames .jpeg to .jpg
        /// </summary>
        public bool Transcode { get; set; }

        #endregion Operations

        /// <summary>
        /// The destination directory - files will be copied or moved there.
        /// </summary>
        public string DestinationDirectory { get; set; }

        public int SelectFiles(string path, bool recursive)
        {
            int count = 0;

            try
            {
                string directory;
                string pattern;

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

                DirectoryInfo di = new DirectoryInfo(directory);
                foreach (var fi in di.EnumerateFiles(pattern, recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly))
                {
                    if (MediaFile.IsSupportedMediaType(fi.Extension))
                    {
                        if (m_selectedFilesHash.Add(fi.FullName.ToLowerInvariant()))
                        {
                            m_selectedFiles.Add(new ProcessFileInfo(fi));
                            m_selectedFilesSize += fi.Length;
                            ++count;
                        }
                    }
                }
            }
            catch (Exception err)
            {
                throw new ArgumentException($"Source '{path}' not found. ({err.Message})", err);
            }

            return count;
        }

        public void PerformOperations()
        {
            m_selectedFilesHash = null; // No longer needed once selection is complete

            // Move or copy the files if a destination directory is specified
            if (DestinationDirectory != null)
            {
                CopyFiles();
            }

            ProcessMediaFiles();

            OnProgressReport(null);
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
                using (var mdf = new MediaFile(fi.Filepath, Path.GetFileName(fi.OriginalFilepath)))
                {
                    mdf.OriginalDateCreated = fi.OriginalDateCreated;
                    mdf.OriginalDateModified = fi.OriginalDateModified;

                    if (SetOrderedNames && mdf.SetOrderedName())
                    {
                        fi.Filepath = mdf.Filepath;
                        OnProgressReport("   Rename to: " + Path.GetFileName(mdf.Filepath));
                    }

                    if (mdf.DeterimineCreationDate())
                    {
                        OnProgressReport($"   Date {mdf.CreationDate.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)} ({mdf.CreationDate.Kind}) from {mdf.CreationDateSource}.");
                    }

                    if (mdf.DetermineTimezone())
                    {
                        OnProgressReport($"   Timezone {mdf.Timezone} from {mdf.TimezoneSource}.");
                    }

                    if (AutoRotate && mdf.Orientation != 1)
                    {
                        OnProgressReport("   Autorotate");
                        mdf.RotateToVertical();
                    }

                    if (Transcode && !mdf.IsPreferredFormat)
                    {
                        OnProgressReport($"   Transcode to {mdf.PreferredFormat}");
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

                    if (!mdf.CommitMetadata())
                    {
                        OnProgressReport("   ERROR: Failed to update metadata. File may be invalid.");
                    }
                }
            }
            catch (Exception err)
            {
                OnProgressReport($"   Error: {err.Message}");
            }
        }

        private void CopyFiles()
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

        private static char[] s_wildcards = new char[] { '*', '?' };

        private static bool HasWildcard(string filename)
        {
            return filename.IndexOfAny(s_wildcards) >= 0;
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
