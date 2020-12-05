using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.Text;
using System.Threading.Tasks;

namespace FMPhotoFinish
{
    class DcimSource : IMediaSource
    {
        const string c_dcimDirectory = "DCIM";
#if DEBUG
        const string c_dcimTestDirectory = "DCIM_Test";
#endif

        // DCF Directories from which files were selected
        // This supports later cleanup of empty directories when move is used.
        List<string> m_dcfDirectories = new List<string>();

        public void RetrieveMediaFiles(SourceConfiguration sourceConfig, IMediaQueue mediaQueue)
        {
            // DCIM requires destination directory
            if (string.IsNullOrEmpty(sourceConfig.DestinationDirectory))
                throw new Exception("DCIM source requires -d destination directory.");

            var queue = SelectDcimFiles(sourceConfig, mediaQueue);
            FileMover.CopyOrMoveFiles(queue, sourceConfig, mediaQueue);
            CleanupDcfDirectories(mediaQueue);
        }

        private List<ProcessFileInfo> SelectDcimFiles(SourceConfiguration sourceConfig, IMediaQueue mediaQueue)
        {
            if (sourceConfig.SelectIncremental) throw new Exception("DCIM source is not compatible with -selectIncremental");
            if (sourceConfig.SelectAfter.HasValue) throw new Exception("DCIM source does not support -selectAfter");

            var sourceFolders = new List<string>();

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

            var queue = new List<ProcessFileInfo>();

            // Add files from each source folder
            foreach (var path in sourceFolders)
            {
                int bookmark = queue.Count;
                mediaQueue.ReportProgress($"Selecting from: {path}");

                DirectoryInfo di = new DirectoryInfo(path);
                foreach (var fi in di.EnumerateFiles())
                {
                    if ((queue.Count % 100) == 0)
                    {
                        mediaQueue.ReportStatus($"Selected: {queue.Count}");
                    }

                    if (MediaFile.IsSupportedMediaType(fi.Extension))
                    {
                        queue.Add(new ProcessFileInfo(fi));
                    }
                }

                if (queue.Count > bookmark) m_dcfDirectories.Add(path);

                mediaQueue.ReportStatus(null);
                mediaQueue.ReportProgress($"   Selected: {queue.Count - bookmark}");
            }

            return queue;
        }

        /// <summary>
        /// Remove empty DCF directories from which files were moved.
        /// </summary>
        private void CleanupDcfDirectories(IMediaQueue mediaQueue)
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
                    mediaQueue.ReportProgress($"Error cleaning up folders on removable drive '{Path.GetPathRoot(directoryName)}': {err.Message}");
                }
            }
        }

#if DEBUG
        /// <summary>
        /// For testing purposes. Copies files from a DCIM_Test directory into the DCIM
        /// directory.
        /// </summary>
        /// <returns>The number of files copied.</returns>
        /// <remarks>
        /// <para>This facilitates testing the <see cref="MoveFiles"/> feature by copying a set
        /// of test files into the DCIM directory before they are moved out.
        /// </para>
        /// <para>For this method to do anything, a removable drive (e.g. USB or SD) must
        /// be present with both DCIM and DCIM_Test directories on it. To be an effective
        /// test - CopyDcimTestFiles must be invoked before SelectDcfFiles.
        /// </para>
        /// </remarks>
        public static int CopyDcimTestFiles()
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
#endif

    }
}
