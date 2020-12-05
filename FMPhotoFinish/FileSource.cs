using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.Text;
using System.Threading.Tasks;

namespace FMPhotoFinish
{
    class FileSource : IMediaSource
    {
        string m_path;
        string m_directory;
        string m_pattern;
        bool m_recursive;
        DateTime m_newestSelection;

        /// <summary>
        /// Select files for processing using directory path with wildcards.
        /// </summary>
        /// <param name="path">The path (including possible wildcards) of files to select.</param>
        /// <param name="recursive">If true, traverse subdirectories for matching files as well.</param>
        /// <remarks>
        /// <para>Wildcards may only appear in the filename, not in the earlier parts of the path.
        /// </para>
        /// </remarks>
        public FileSource(string sourcePattern, bool recursive)
        {
            // Throws an exception if there's an error in the pattern.
            ParseSelectFilesPath(sourcePattern, out m_directory, out m_pattern);
            m_path = $"{m_directory}\\{m_pattern}";
            m_recursive = recursive;
        }

        public void RetrieveMediaFiles(SourceConfiguration sourceConfig, IMediaQueue mediaQueue)
        {
            var queue = SelectFiles(sourceConfig, mediaQueue);

            // if a destination directory was specified, copy or move the files
            if (sourceConfig.DestinationDirectory != null)
            {
                FileMover.CopyOrMoveFiles(queue, sourceConfig, mediaQueue);
            }

            // Else, simply put them in the mediaQueue
            else
            {
                FileMover.EnqueueFiles(queue, mediaQueue);
            }

            // Save the bookmark
            if (m_newestSelection > DateTime.MinValue)
            {
                // Only sets a bookmark if incremental is on.
                sourceConfig.SetBookmark(m_path, m_newestSelection);
            }
        }

        private List<ProcessFileInfo> SelectFiles(SourceConfiguration sourceConfig, IMediaQueue mediaQueue)
        {
            mediaQueue.ReportProgress($"Selecting {(m_recursive ? "from" : "tree")}: {m_path}");

            // Determine the "after" threshold from SelectAfter and SelectIncremental
            var after = sourceConfig.GetBookmarkOrAfter(m_path);

            m_newestSelection = after ?? DateTime.MinValue;
            var queue = new List<ProcessFileInfo>();
            int skippedFiles = 0;

            try
            {
                DirectoryInfo di = new DirectoryInfo(m_directory);
                foreach (var fi in di.EnumerateFiles(m_pattern, m_recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly))
                {
                    if (((queue.Count + skippedFiles) % 100) == 0)
                    {
                        string message = (skippedFiles == 0)
                            ? $"Selected: {queue.Count}"
                            : $"Selected: {queue.Count} Not Selected: {skippedFiles}";
                        mediaQueue.ReportStatus(message);
                    }

                    if (MediaFile.IsSupportedMediaType(fi.Extension))
                    {
                        // Limit date window of files to be selected
                        if (after.HasValue)
                        {
                            var date = MediaFile.GetBookmarkDate(fi.FullName);
                            if (!date.HasValue || date.Value <= after.Value)
                            {
                                ++skippedFiles;
                                continue;
                            }

                            // For this operation, we do everything in localtime because photo
                            // DateTaken metadata is in localtime.
                            // This is after-the-fact but since it's a debugging test that's OK.
                            Debug.Assert(date.Value.Kind == DateTimeKind.Local);

                            if (m_newestSelection < date.Value)
                                m_newestSelection = date.Value;
                        }

                        queue.Add(new ProcessFileInfo(fi));
                    }
                }
            }
            catch (Exception err)
            {
                throw new ArgumentException($"Source '{m_path}' not found. ({err.Message})", err);
            }
            mediaQueue.ReportStatus(null);
            mediaQueue.ReportProgress(skippedFiles == 0
                ? $"   Selected: {queue.Count}"
                : $"   Selected: {queue.Count} Not Selected: {skippedFiles}");

            // If SelectIncremental, report the new bookmark
            if (sourceConfig.SelectIncremental && queue.Count > 0)
            {
                Debug.Assert(m_newestSelection > DateTime.MinValue);
                mediaQueue.ReportProgress($"   Newest: {m_newestSelection:yyyy'-'MM'-'dd' 'HH':'mm':'ss}");
            }

            return queue;
        }



        private static void ParseSelectFilesPath(string path, out string directory, out string pattern)
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

        private static char[] s_wildcards = new char[] { '*', '?' };
        private static bool HasWildcard(string filename)
        {
            return filename.IndexOfAny(s_wildcards) >= 0;
        }
    }
}
