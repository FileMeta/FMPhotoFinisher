using System;
using System.IO;
using System.Text;

//Next Steps
/* Done:
 *  -d
 *  -autorot
 *  -transcode video
 *  -orderednames
 *  -transcode audio
 *  -log
 *  determine timezone
 *  preserve metadata on transcode and store updated metadata
 * Next:
 *  -setTimezone
 *  -changeTimezone
 *  -sort
 *  -st
 *  -sDCF
 *  add uuid for each image
 * 
 * Design Principles:
 *  * No command-line option required to preserve, transfer, or add information. For example,
 *    copying the timezone from a makernote field to a custom field should be done regardless
 *    of command-line options. But, changing date_created from local time to UTC requires the
 *    "-datefixup" option.
 * 
 * Framework: Process the queue
 *   Retrieve metadata
 *   Rotate images
 *   Transcode
 *   Update metadata
 */

namespace FMPhotoFinish
{
    class Program
    {
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
  -orderednames    Canon cameras prefix photos with ""IMG_"" and videos with
                   ""MVI_"". This option renames videos to use the ""IMG_""
                   prefix thereby having them show in order with the
                   associated photos.

  -autorot         Using the 'orientation' metadata flag, auto rotate images
                   to their vertical position and clear the orientation flag.

  -transcode       Transcode video and audio files to the preferred format
                   which is .mp4 for video and .m4a for audio. Also renames
                   .jpeg files to .jpg.

Other Options:

  -h               Print this help text and exit (ignoring all other
                   commands).

  -log             Log all operations. The file will be named
                   ""<date> <time> FMPhotoFinish Log.txt"". If -d is
                   specified then the log will be placed in that folder.
                   Otherwise, the log is placed in a ""logs"" directory under
                   the user's ""documents"" folder. (Typically
                   ""C:\users\<username>\documents\logs"".)

  -w               Wait for the user to press a key before exiting. This is
                   handy when running from a shortcut with no console or when
                   running under the debugger.
";

        const string c_logsFolder = "logs";
        const string c_logFileSuffix = "FMPhotoFinish Log.txt";

        static bool s_commandLineError;
        static bool s_showSyntax;
        static bool s_waitBeforeExit;
        static bool s_log;
        static TextWriter s_logWriter;

        static void Main(string[] args)
        {
            try
            {
                var photoFinisher = new PhotoFinisher();
                photoFinisher.ProgressReported += ReportProgress;
                photoFinisher.StatusReported += ReportStatus;
                ParseCommandLine(args, photoFinisher);

                if (s_commandLineError)
                {
                    // Do nothing
                }
                else if (s_showSyntax)
                {
                    Console.WriteLine(c_Syntax);
                }
                else
                {
                    // Prepare logfile
                    if (s_log)
                    {
                        string logDir = photoFinisher.DestinationDirectory;
                        if (string.IsNullOrEmpty(logDir))
                        {
                            logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Personal), c_logsFolder);
                            if (!Directory.Exists(logDir))
                            {
                                Directory.CreateDirectory(logDir);
                            }
                        }

                        var now = DateTime.Now;
                        var logfilename = Path.Combine(logDir,
                            $"{now.ToString("yyyy-MM-dd hhmmss")} {c_logFileSuffix}");

                        s_logWriter = new StreamWriter(logfilename, false, Encoding.UTF8);

                        // Metadata in MicroYAML format
                        s_logWriter.WriteLine("---");
                        s_logWriter.WriteLine($"date: {now.ToString("yyyy-MM-dd hh:mm:ss")}");
                        s_logWriter.WriteLine("title: FMPhotoFinish Log");
                        s_logWriter.WriteLine($"commandline: {Environment.CommandLine}");
                        s_logWriter.WriteLine("...");
                    }

                    // Do the work.
                    photoFinisher.PerformOperations();
                }
            }
            catch (Exception err)
            {
#if DEBUG
                Console.WriteLine(err.ToString());
#else
                Console.WriteLine(err.Message);
#endif
                if (s_logWriter != null)
                {
                    s_logWriter.WriteLine(err.ToString());
                }
            }
            finally
            {
                if (s_logWriter != null)
                {
                    s_logWriter.Dispose();
                    s_logWriter = null;
                }
            }

            if (s_waitBeforeExit)
            {
                Console.WriteLine();
                Console.WriteLine("Press any key to exit.");
                Console.ReadKey();
            }
        }

        static void ParseCommandLine(string[] args, PhotoFinisher photoFinisher)
        {
            try
            {
                if (args.Length == 0)
                {
                    s_showSyntax = true;
                    return;
                }

                for (int i = 0; i < args.Length; ++i)
                {
                    switch (args[i].ToLowerInvariant())
                    {
                        case "-h":
                        case "-?":
                            s_showSyntax = true;
                            break;

                        case "-s":
                            {
                                ++i;
                                int n = photoFinisher.SelectFiles(args[i], false);
                                Console.WriteLine($"Selected {n} from \"{args[i]}\"");
                            }
                            break;

                        case "-d":
                            ++i;
                            {
                                string dst = args[i];
                                if (!Directory.Exists(dst))
                                {
                                    Console.WriteLine($"Destination folder '{dst}' does not exist.");
                                    s_commandLineError = true;
                                    break;
                                }
                                photoFinisher.DestinationDirectory = Path.GetFullPath(dst);
                            }
                            break;

                        case "-w":
                            s_waitBeforeExit = true;
                            break;

                        case "-sort":
                            photoFinisher.AutoSort = true;
                            break;

                        case "-move":
                            photoFinisher.Move = true;
                            break;

                        case "-transcode":
                            photoFinisher.Transcode = true;
                            break;

                        case "-orderednames":
                            photoFinisher.SetOrderedNames = true;
                            break;

                        case "-autorot":
                            photoFinisher.AutoRotate = true;
                            break;

                        case "-log":
                            s_log = true;
                            break;

                        default:
                            Console.WriteLine($"Command-line syntax error: '{args[i]}' is not a recognized command.");
                            Console.WriteLine("Use '-h' for syntax help");
                            s_commandLineError = true;
                            break;
                    }
                }
            }
            catch (Exception err)
            {
                Console.WriteLine(err.Message);
                s_commandLineError = true;
            }

        } // ParseCommandLine

        static void ReportProgress(object obj, ProgressEventArgs eventArgs)
        {
            Console.Out.WriteLine(eventArgs.Message ?? string.Empty);
            if (s_logWriter != null)
            {
                s_logWriter.WriteLine(eventArgs.Message ?? string.Empty);
            }
        }
        
        static void ReportStatus(object obj, ProgressEventArgs eventArgs)
        {
            if (string.IsNullOrEmpty(eventArgs.Message))
            {
                Console.Error.WriteLine();
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.Error.Write(eventArgs.Message);
                Console.Error.Write('\r');
                Console.ForegroundColor = ConsoleColor.White;
            }
        }

    } // Class Program

    static class FormatExtensions
    {
        public static string FmtCustom(this TimeSpan ts)
        {
            var sb = new StringBuilder();
            int hours = (int)ts.TotalHours;
            if (hours > 0)
            {
                sb.Append(hours.ToString());
                sb.Append(':');
            }
            sb.Append(ts.Minutes.ToString("d2"));
            sb.Append(':');
            sb.Append(ts.Seconds.ToString("d2"));
            return sb.ToString();
        }
    }
}