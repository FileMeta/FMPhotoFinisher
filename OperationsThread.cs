using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.IO;
using System.Text;
using System.Threading.Tasks;

//Next Steps
/* Parse -d parameter
 * Copy files according to -d and update the files in the queue
 *   Ensure that files with the same name are handled properly (add a numeric suffix)
 * Process the queue
 *   Retrieve metadata
 *   Rotate images
 *   Update metadata
 *   Note: Use a prefix character on keywords that are extended metadata (e.g. *, &, :, !, ?)
 */

namespace FMPhotoFinisher
{
    class OperationsThread
    {
// Column 78                                                                 |
        const string c_Syntax =
@"
FMPhotoFinisher processes photos, videos, and audio files normalizing
metadata, rotating photos to vertical, transcoding video, and so forth. It
facilitates collecting media from source devices like video cameras, smart
phones and audio recorders.

Command-Line Syntax:
  FMPhotoFinisher <source> [destination] <operations>

Supported File Types:
  Only files with the following extensions are supported. All others are
  ignored:
    .jpg .jpeg       JPEG Images
    .mp4             MPEG 4 Video
    .m4a             MPEG 4 Audio
    .mp3             MP3 Audio    

    Other video formats; must be transcoded to .mp4 format to gain metadata
    support;
    .avi .mov .mpg .mpeg

    Other audio formats; must be transcoded to .m4a or .mp3 to gain metadata
    support.
    .wav

Source:
  The source selection parameters indicate the files to be processed.
  Regardless of selection mode or wildcards, only files with the extensions
  listed above will be processed; all others are ignored. Multiple selection
  parameters may be included. If more than one selection happens to include
  the same file, the file will only be processed once.

  -s <path>        Select files from the specified path. If the path
                   specifies a directory then all files in the directory are
                   included (so long as they have a supported filename
                   extension). The path may include wildcards.
  
  -st <path>       Select files in the specified subdirectory tree. This is
                   similar to -s except that all files in subdirectories of
                   the specified path are also included.

  -sDCF            Select files from all removable DCF (Design rule for
                   Camera File system) devices. This is the best option for
                   retrieving from digital cameras or memory cards from
                   cameras.

Destination:
  If no destination is specified, then the updates are made in-place.

  -d <path>        A path to the folder where the processed files should be
                   placed.

  -sort            Auto-sort the transferred files into a directory tree
                   rooted in the folder specified by -d. The tree hierarchy
                   is year/month/day.

  -move            Move the files from the source to the destination. If this
                   parameter is not included then the unprocessed original
                   files are left at the source.

Operations:
  -autorot         Using the 'orientation' metadata flag, auto rotate images
                   to their vertical position and clear the orientation flag.

  -transcode       Transcode video and audio files to the preferred format
                   which is .mp4 for video and .m4a for audio. Also renames
                   .jpeg files to .jpg.

  -datefixup       If the ""Date Taken"" field (JPEG images) or ""Media Created""
                   field (MP4 and M4A video and audio files) is missing, fill
                   it in with the best available information including the
                   date created, date modifed, and dates from neighboring
                   files in the selection. In particular, this is effective
                   for filling in the date for video files when cameras don't
                   supply that information. Also fills in timezone tag if it
                   is not present.

Other Options:

  -h               Print this help text and exit (ignoring all other
                   commands).

  -log <filename>  Log all operations to the specified file. If the file
                   exists then the new operations will be appended to the end
                   of the file.
";
// Column 78                                                                 |

        int m_started;  // Actually used as a bool but Interlocked works better with an int.
        Thread m_thread;
        MainWindow m_mainWindow;

        // Command-line operations
        bool m_showSyntax;
        bool m_commandLineError;
        List<ProcessFileInfo> m_selectedFiles = new List<ProcessFileInfo>();
        HashSet<string> m_selectedFilesHash = new HashSet<string>();
        long m_selectedFilesSize;
        bool m_autoRotate;
        string m_dstDirectory;
        bool m_sort;
        bool m_move;
        bool m_transcode;
        bool m_dateFixup;

        public OperationsThread(MainWindow mainWindow)
        {
            m_thread = new Thread(ThreadMain);
            m_mainWindow = mainWindow;
        }

        public void Start()
        {
            int started = Interlocked.CompareExchange(ref m_started, 1, 0);
            if (started == 0)
            {
                m_thread.Start();
            }
        }

        private void ThreadMain()
        {
            GC.KeepAlive(this); // Paranoid - probably not necessary as the thread has a reference to this and the system has a reference to the thread.

            try
            {
                using (var exifTool = new ExifToolWrapper.ExifTool())
                {
                    var dict = new Dictionary<string, string>();
                    exifTool.GetProperties("F:\\SampleData\\PhotoSource\\IMG_4732.jpg", dict);
                    foreach (var pair in dict)
                    {
                        m_mainWindow.WriteLine($"{pair.Key}={pair.Value}");
                    }
                }

                ParseCommandLine();
                if (m_commandLineError) return;
                if (m_showSyntax)
                {
                    m_mainWindow.WriteLine(c_Syntax);
                    return;
                }

                // Move or copy the files if a destination directory is specified
                if (m_dstDirectory != null)
                {
                    CopyFiles();
                }

                ProcessMediaFiles();

                m_mainWindow.WriteLine();
                m_mainWindow.WriteLine("All operations complete!");
                m_mainWindow.SetProgress(string.Empty);
            }
            catch (Exception err)
            {
                m_mainWindow.WriteLine(err.ToString());
            }
        }

        void ProcessMediaFiles()
        {
            // We index through the files because we may need the index to match metadata
            // on a preceding or succeeding file
            int count = m_selectedFiles.Count;
            for (int i=0; i<count; ++i)
            {
                var fi = m_selectedFiles[i];
                ProcessMediaFile(fi, i);
            }
        }

        void ProcessMediaFile(ProcessFileInfo fi, int index)
        {
            bool ann = false;

            using (var mdf = new MediaFile(fi.Filepath))
            {
                if (m_autoRotate && mdf.Orientation != 1)
                {
                    AnnounceFile(fi, ref ann);
                    m_mainWindow.WriteLine("   Autorotate");
                    mdf.RotateToVertical();
                }

                if (m_transcode && !mdf.IsPreferredFormat)
                {
                    AnnounceFile(fi, ref ann);
                    m_mainWindow.WriteLine($"   Transcode to {mdf.PreferredFormat}");
                    mdf.TranscodeToPreferredFormat();
                }

                if (m_dateFixup)
                {
                    // Fill in date taken if not present
                    // Fill in timezone if not present

                }
            }
        }

        void AnnounceFile(ProcessFileInfo fi, ref bool announced)
        {
            if (!announced)
            {
                m_mainWindow.WriteLine(fi.Filepath);
                announced = true;
            }
        }

        void ParseCommandLine()
        {
            string[] args = Environment.GetCommandLineArgs();
            if (args.Length == 1)
            {
                m_showSyntax = true;
                return;
            }

            for (int i=1; i<args.Length; ++i)
            {
                switch (args[i].ToLowerInvariant())
                {
                    case "-h":
                        m_showSyntax = true;
                        break;

                    case "-s":
                        ++i;
                        SelectFiles(args[i], false);
                        break;

                    case "-d":
                        ++i;
                        {
                            string dst = args[i];
                            if (!Directory.Exists(dst))
                            {
                                m_mainWindow.WriteLine($"Command-line syntax error: Destination folder '{dst}' does not exist.");
                                m_commandLineError = true;
                                break;
                            }
                            m_dstDirectory = Path.GetFullPath(dst);
                        }
                        break;

                    case "-sort":
                        m_sort = true;
                        break;

                    case "-move":
                        m_move = true;
                        break;

                    case "-transcode":
                        m_transcode = true;
                        break;

                    case "-autorot":
                        m_autoRotate = true;
                        break;

                    case "-datefixup":
                        m_dateFixup = true;
                        break;

                    default:
                        m_mainWindow.WriteLine($"Command-line syntax error: '{args[i]}' is not a recognized command.");
                        m_mainWindow.WriteLine();
                        m_showSyntax = true;
                        break;
                }
            }

            m_selectedFilesHash = null; // No longer needed once selection is complete
        }

        private void SelectFiles(string path, bool recursive)
        {
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
                        m_mainWindow.WriteLine($"Source '{path}' is invalid. Wildcards not allowed in directory name.");
                        m_commandLineError = true;
                        return;
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
                        m_mainWindow.WriteLine($"Source '{path}' does not exist.");
                        m_commandLineError = true;
                        return;
                    }
                }

                int count = 0;
                int dupCount = 0;
                DirectoryInfo di = new DirectoryInfo(directory);
                foreach(var fi in di.EnumerateFiles(pattern, recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly))
                {
                    if (MediaFile.IsSupportedMediaType(fi.Extension))
                    {
                        if (m_selectedFilesHash.Add(fi.FullName.ToLowerInvariant()))
                        {
                            m_selectedFiles.Add(new ProcessFileInfo(fi));
                            m_selectedFilesSize += fi.Length;
                            ++count;
                        }
                        else
                        {
                            ++dupCount;
                        }
                    }
                }

                if (count == 0)
                {
                    if (dupCount == 0)
                    {
                        m_mainWindow.WriteLine($"No media files found at source '{path}'.");
                        m_commandLineError = true;
                    }
                    else
                    {
                        m_mainWindow.WriteLine($"All media files found at source '{path}' are duplicates.");
                    }
                }
            }
            catch (Exception err)
            {
                m_mainWindow.WriteLine($"Source '{path}' not found. ({err.Message})");
                m_commandLineError = true;
            }

        }

        private void CopyFiles()
        {
            string verb = m_move ? "Moving" : "Copying";
            string dstDirectory = m_dstDirectory;
            // TODO: If -sort then add a temporary directory.

            m_mainWindow.WriteLine($"{verb} media files to working folder.");

            uint startTicks = (uint)Environment.TickCount;
            long bytesCopied = 0;

            int n = 0;
            foreach (var fi in m_selectedFiles)
            {
                if (bytesCopied == 0)
                {
                    m_mainWindow.SetProgress($"{verb} file {n + 1} of {m_selectedFiles.Count}");
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

                    m_mainWindow.SetProgress($"{verb} file {n + 1} of {m_selectedFiles.Count}. Time remaining: {remain.FmtRemain()} MBps: {(bps/(1024*1024)):#,###.###}");
                }

                string dstFilepath = Path.Combine(m_dstDirectory, Path.GetFileName(fi.OriginalFilepath));
                MediaFile.MakeFilepathUnique(ref dstFilepath);

                if (m_move)
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
        }

        private static char[] s_wildcards = new char[] { '*', '?' };

        private static bool HasWildcard(string filename)
        {
            return filename.IndexOfAny(s_wildcards) >= 0;
        }

        private void OldCode()
        {
            /*
            m_mainWindow.OutputWrite("Operations complete. Exiting in 5 seconds.");
            Thread.Sleep(5000);
            m_mainWindow.Dispatcher.BeginInvokeShutdown(System.Windows.Threading.DispatcherPriority.Normal);
            */


            /*
            // Get the destination folder name from the command line
            string destFolder;
            {
                string[] args = Environment.GetCommandLineArgs();
                destFolder = (args.Length == 2) ? args[1] : string.Empty;
            }

            bool showSyntax = false;
            if (destFolder.Length == 0
                || string.Equals(destFolder, "-h", StringComparison.OrdinalIgnoreCase)
                || string.Equals(destFolder, "/h", StringComparison.OrdinalIgnoreCase))
            {
                showSyntax = true;
            }
            else if (!Directory.Exists(destFolder))
            {
                m_mainWindow.OutputWrite("Destination folder '{0}' does not exist.\r\n", destFolder);
                showSyntax = true;
            }
            else
            {
                destFolder = Path.GetFullPath(destFolder);
            }
            if (showSyntax)
            {
                m_mainWindow.OutputWrite("Command-Line Syntax: FMCollectFromCamera <destination folder path>\r\n");
                return;
            }

            m_mainWindow.OutputWrite("Collecting images from cameras and cards to '{0}'.\r\n", destFolder);

            IntPtr hwndOwner = m_mainWindow.GetWindowHandle();

            // Process each removable drive
            foreach (DriveInfo drv in DriveInfo.GetDrives())
            {
                if (drv.IsReady && drv.DriveType == DriveType.Removable)
                {
                    // File system structure is according to JEITA "Design rule for Camera File System (DCF) which is JEITA specification CP-3461
                    // See if the DCIM folder exists
                    DirectoryInfo dcim = new DirectoryInfo(Path.Combine(drv.RootDirectory.FullName, "DCIM"));
                    if (dcim.Exists)
                    {
                        List<string> sourceFolders = new List<string>();
                        List<string> sourcePaths = new List<string>();

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

                                    foreach (string ext in s_mediaExtensions)
                                    {
                                        if (di.EnumerateFiles(string.Concat("*", ext)).Any())
                                        {
                                            sourcePaths.Add(Path.Combine(di.FullName, string.Concat("*", ext)));
                                        }
                                    }
                                }
                            }
                        }

                        if (sourcePaths.Count > 0)
                        {
                            // Write the names out
                            foreach (string path in sourcePaths)
                            {
                                m_mainWindow.OutputWrite(path);
                                m_mainWindow.OutputWrite("\r\n");
                            }

                            // Perform the move
                            WinShell.SHFileOperation.MoveFiles(hwndOwner, "Collecting Photos from Camera", sourcePaths, destFolder);

                            // Clean up the folders
                            try
                            {
                                foreach (string folderName in sourceFolders)
                                {
                                    DirectoryInfo di = new DirectoryInfo(folderName);

                                    // Get rid of thumbnails unless a matching file still exists
                                    // (matching files have to be of a type we don't yet recognize)
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
                            }
                            catch (Exception err)
                            {
                                // Report errors during cleanup but proceed with other files.
                                m_mainWindow.OutputWrite("Error cleaning up folders on drive '{0}'.\r\n{0}\r\n", drv.Name, err.Message);
                            }
                        }

                    } // If DCIM exists
                } // If drive is ready and removable
            } // for each drive

            */
        } // Function ThreadMain

    } // class OperationsThread

    static class FormatExtensions
    {
        public static string FmtRemain(this TimeSpan ts)
        {
            var sb = new StringBuilder();
            int hours = (int)ts.TotalHours;
            if (hours > 0)
            {
                sb.Append(hours.ToString());
                sb.Append(':');
            }
            if (hours > 0 || ts.Minutes > 0)
            {
                sb.Append(ts.Minutes.ToString("d2"));
                sb.Append(':');
            }
            sb.Append(ts.Seconds.ToString("d2"));
            return sb.ToString();
        }
    }
}
