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
 *  -setTimezone
 *  -changeTimezone
 *  -setDate
 *  -setUUID
 *  -saveOriginalFn
 *  -sort
 *  -sDCIM
 *  -updateFsDate
 *  -st
 *  -shiftDate
 * Next:
 * --- Ready to begin using
 *  Test with no options - should just list values (potential conflict with design principles - figure this out)
 *  Tabular output format option (for later analytics)
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

  -sDCIM           Select files from all removable DCF (Design rule for
                   Camera File system) devices. These are removable drives
                   with a DCIM folder in the root - typically used by digital
                   cameras. This is the best option for retrieving from
                   digital cameras or memory cards from cameras.

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

  -saveOriginalFn  Save the original filename in a custom metaTag stored
                   in the comments property. If the property already exists,
                   it is preserved.

  -setUUID         Set a unique UUID as a custom metaTag stored in the
                   comments property. If the property already exists, it is
                   preserved.

  -setDate <dateTime>  Set the dateCreated/dateTaken value in the metadata
                   to the specified date and time. (See details on format
                   of the date below.)

  -shiftDate <+|-><[dd:]hh:mm:ss>
  -shiftDate <target dateTime> <original dateTime> 
                   Shifts the date/time of files by the specified amount. The
                   amount to shift can be specified in either of two ways.
                   In the first case, the offset is given by a number of
                   days, hours, minutes and seconds. In the second case, the
                   amount is specified by taking the difference between two
                   dateTimes. (See details on both forms below.)

  -setTimezone <tz> 
                   Set the timezone to the specified value keeping the
                   local time the same. (See details on timezone below.)

  -changeTimezone <tz>  Change the timezone to the specified value keeping
                   the UTC time the same. (See details on timezone below.)

  -updateFsDate    Update the file system dateCreated to match the metadata.
                   Note that the file system dateModified value

Other Options:

  -h               Print this help text and exit (ignoring all other
                   commands).

  -listTimezones   List all timezone IDs and associated offsets.

  -log             Log all operations. The file will be named
                   ""<date> <time> FMPhotoFinish Log.txt"". If -d is
                   specified then the log will be placed in that folder.
                   Otherwise, the log is placed in a ""logs"" directory under
                   the user's ""documents"" folder. (Typically
                   ""C:\users\<username>\documents\logs"".)

  -w               Wait for the user to press a key before exiting. This is
                   handy when running from a shortcut with no console or when
                   running under the debugger.

SetDate:
  The date-time value should be in W3CDTF format (see
  https://www.w3.org/TR/NOTE-datetime). Here is an example:
    2018-12-20T17:55:24-07:00

  This means, 20 December 2018 at 5:55:24 PM Mountain Standard Time (UTC-7)

  You can leave off the seconds or minutes. For example:
    2018-12-20T17:55-07:00

  If you leave off the timezone then the current timezone setting of the
  computer will be used.
    2018-12-20T17:55

  You can specify just the date. If only the date is specified, the time is
  set to 12:00 noon (middle of the day) to avoid having the date shift around
  when just one timezone away.
    2018-12-20

  You can even specify just the month or year. When setting anything less
  precise than the hour and minute, the custom ""datePrecision"" metadata
  tag will be set to the number of significant digits in the date and time
  as follows: 4=year, 6=month, 8=day, 10=hour, 12=minute, 14=second,
  17=millisecond precision.

  Instead of putting appending the timezone to the date-time value, You can
  set the timezone separately using the -setTimezone argument. You can even
  set the time in one timezone and then change it to another by combining
  -setDate with the -changeTimezone argument.

  Regardless of their order on the command-line, arguments are processed in
  this order: -setDate -setTimezone -changeTimezone.

ShiftDate:
  The amount to shift can be specified in either of two ways; In the first
  case, the offset is given by a number of days, hours, minutes and seconds.
  In the second case, the amount is specified by taking the difference
  between two dateTimes.

  An offset MUST be preceded by a sign (plus or minus) and follows this
  format:

  {+|-}{ d | d.hh:mm[:ss[.ff]] | hh:mm[:ss[.ff]] }

  Elements in square brackets ([ and ]) are optional. One selection from the
  list of alternatives enclosed in braces ({ and }) and separated by vertical
  bars (|) is required. d=days, hh=hours, mm=minutes, ss=seconds, ff=fractions
  of a second. Here are examples:

    -shiftDate +3:00:00
       Move forward 3 hours. For example, 3:25 becomes 3:25.

    -shiftDate -2:15:00
       Move backward 2 hours and 15 minutes. For example, 3:25 becomes 1:10.

    -shiftDate +2.01:00:00
       Move forward two days and one hour.

    -shiftDate -1
       Move backward one day.

  The second way to specify the offset is to give two dateTimes, a target and
  an original. Typically this is used when the time was set incorrectly on
  a camera. You pick the time that's recorded on one of the photos (the
  original) and then specify what that time should be (the target).
  FMPhotoFinish takes the difference between the two times and shifts all
  times in the set by that same amount.

  Dates should be specified in W3CDTF format. See the SetDate information
  above for details. Here are examples.

    -shiftDate 2016-07-15 2015-07-10
       Move forward one year and five days. A photo with the date of
       July 10, 2015 would be changed to July 15, 2016 (while keeping the
       time of day the same).

    -shiftDate 2016-07-15T10:00 2017-07-15T11:00
       Move backward one year and one day. A photo with the date of
       July 10, 2017 at 8:00 would be changed to July 10, 2016 at 7:00.

Timezones:
  The -setTimezone and -changeTimezone options are similar with an important
  difference.

  -setTimezone keeps local time the same while changing the timezone. So, if
  a photo was taken at 9:30 AM EST, which is 14:30 UTC, setting timezone to
  Pacific will result in 9:30 AM PST which is 17:30 UTC. This is useful when
  the camera was set to the right local time but didn't have the correct
  timezone set or doesn't even have a timezone setting.

  -changeTimezone keeps UTC the same while changing the timezone. So, if a
  photo was taken at 9:30 AM EST, which is 14:30 UTC, setting timezone to
  pacific will result in 6:30 AM PSt, which remains 14:30 UTC. This is useful
  when the time was set correctly at home and then you travelled to another
  timezone without resetting the time on the camera.

  Both operations require a timezone ID or offset. To get a list of
  timezone IDs, use the ""-listTimezones"" command-line option. A timezone
  offset is in the form of hours and minutes before or after UTC. Examples
  are ""-05:00"" (Eastern), ""+01:00"" (European), ""+09:30"" (Australian
  Central). If minutes are zero they can be left off. E.g. ""-05"", ""+03"".

  Internally, JPEG photos store times in local time while MP4 videos store
  time in UTC. FMPhotoFinish adds a custom timezone field so that times can
  be determined consistently between the two even when changing timezone
  settings on the computer. Many cameras don't have a timezone setting and
  so they store the time for MP4 videos in local time even though the spec
  says it should be UTC. Upon initial processing, FMPhotoFinish sets the
  timezone on those files to ""0"" which means that the time should be
  interpreted as local even though the specification says that it's UTC.

  If the timezone is unknown (set to ""0"") then you using -changeTimeZone
  will fail with an error message because the starting timezone is unknown.
  Instead, first use -setTimeZone and then -changeTimeZone. Both options
  may be used in one FMPhotoFinish operation or you can do them in two
  consecutive operations.
";

        const string c_logsFolder = "logs";
        const string c_logFileSuffix = "FMPhotoFinish Log.txt";

        static bool s_commandLineError;
        static bool s_showSyntax;
        static bool s_waitBeforeExit;
        static bool s_log;
        static bool s_listTimezones;
        static TextWriter s_logWriter;

        static void Main(string[] args)
        {
            try
            {
                var photoFinisher = new PhotoFinisher();
                photoFinisher.ProgressReported += ReportProgress;
                photoFinisher.StatusReported += ReportStatus;
                ParseCommandLine(args, photoFinisher);

                PerformOperations(photoFinisher);
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
                                if (i >= args.Length)
                                {
                                    Console.WriteLine($"Expected value for -s command-line argument.");
                                    s_commandLineError = true;
                                    break;
                                }

                                Console.WriteLine($"Selecting from: {args[i]}");
                                int n = photoFinisher.SelectFiles(args[i], false);
                                Console.WriteLine($"   {n} media files selected.");
                            }
                            break;

                        case "-st":
                            {
                                ++i;
                                if (i >= args.Length)
                                {
                                    Console.WriteLine($"Expected value for -st command-line argument.");
                                    s_commandLineError = true;
                                    break;
                                }

                                Console.WriteLine($"Selecting tree: \"{args[i]}\"");
                                int n = photoFinisher.SelectFiles(args[i], true);
                                Console.WriteLine($"   {n} media files selected.");
                            }
                            break;

                        case "-copydcim":
                            {
                                Console.WriteLine("Copying test files.");
                                int n = photoFinisher.CopyDcimTestFiles();
                                Console.WriteLine($"Copied {n} files from DCIM_Test to DCIM.");
                            }
                            break;

                        case "-sdcim":
                            {
                                int n = photoFinisher.SelectDcimFiles();
                                Console.WriteLine($"Selected {n} from removable camera devices.");
                            }
                            break;

                        case "-d":
                            {
                                ++i;
                                if (i >= args.Length)
                                {
                                    Console.WriteLine($"Expected value for -d command-line argument.");
                                    s_commandLineError = true;
                                    break;
                                }

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

                        case "-saveoriginalfn":
                        case "-saveoriginalfilename":
                            photoFinisher.SaveOriginalFilaname = true;
                            break;

                        case "-setuuid":
                            photoFinisher.SetUuid = true;
                            break;

                        case "-autorot":
                            photoFinisher.AutoRotate = true;
                            break;

                        case "-setdate":
                            {
                                ++i;
                                if (i >= args.Length)
                                {
                                    Console.WriteLine($"Expected value for -setDate command-line argument.");
                                    s_commandLineError = true;
                                    break;
                                }

                                FileMeta.DateTag dt;
                                if (!FileMeta.DateTag.TryParse(args[i], out dt))
                                {
                                    Console.WriteLine($"Invalid value for -setDate '{args[i]}'.");
                                    s_commandLineError = true;
                                    break;
                                }

                                // Default to the local timezone if needed
                                dt = dt.ResolveTimeZone(TimeZoneInfo.Local);

                                photoFinisher.SetDateTo = dt;
                            }
                            break;

                        case "-shiftdate":
                            {
                                ++i;
                                if (i >= args.Length)
                                {
                                    Console.WriteLine($"Expected value for -shiftDate command-line argument.");
                                    s_commandLineError = true;
                                    break;
                                }

                                // If the next argument starts with a sign, then the shift amount is a simple timespan.
                                if (args[i][0] == '+' || args[i][0] == '-')
                                {
                                    string s = (args[i][0] == '+') ? args[i].Substring(1) : args[i];
                                    TimeSpan ts;
                                    if (!TimeSpan.TryParse(s, System.Globalization.CultureInfo.InvariantCulture, out ts))
                                    {
                                        Console.WriteLine($"Invalid value for -shiftDate '{args[i]}'.");
                                        s_commandLineError = true;
                                        break;
                                    }

                                    photoFinisher.ShiftDateBy = ts;
                                }

                                // Else, shift amount is the difference of two times
                                else
                                {
                                    if (i+1 >= args.Length)
                                    {
                                        Console.WriteLine($"Expected two values for -shiftDate command-line argument.");
                                        s_commandLineError = true;
                                        break;
                                    }

                                    FileMeta.DateTag dtTarget;
                                    if (!FileMeta.DateTag.TryParse(args[i], out dtTarget))
                                    {
                                        Console.WriteLine($"Invalid value for -shiftDate '{args[i]}'.");
                                        s_commandLineError = true;
                                        break;
                                    }

                                    ++i;
                                    FileMeta.DateTag dtSource;
                                    if (!FileMeta.DateTag.TryParse(args[i], out dtSource))
                                    {
                                        Console.WriteLine($"Invalid value for -shiftDate '{args[i]}'.");
                                        s_commandLineError = true;
                                        break;
                                    }

                                    // Resolve timezone if it was ambiguous.
                                    dtTarget.ResolveTimeZone(TimeZoneInfo.Local);
                                    dtSource.ResolveTimeZone(TimeZoneInfo.Local);

                                    // For whatever reason, they might have used different timezones. Shift both to Utc before taking difference.
                                    photoFinisher.ShiftDateBy = dtTarget.ToUtc().Subtract(dtSource.ToUtc());
                                }
                            }
                            break;

                        case "-settimezone":
                            {
                                ++i;
                                if (i >= args.Length)
                                {
                                    Console.WriteLine($"Expected value for -setTimezone command-line argument.");
                                    s_commandLineError = true;
                                    break;
                                }

                                var tzi = TimeZoneParser.ParseTimeZoneId(args[i]);
                                if (tzi == null)
                                {
                                    Console.WriteLine($"Invalid value for -setTimezone '{args[i]}'. Use '-listTimezones' option to find valid values.");
                                    s_commandLineError = true;
                                    break;
                                }
                                photoFinisher.SetTimezoneTo = tzi;
                            }
                            break;

                        case "-changetimezone":
                            {
                                ++i;
                                if (i >= args.Length)
                                {
                                    Console.WriteLine($"Expected value for -changeTimezone command-line argument.");
                                    s_commandLineError = true;
                                    break;
                                }

                                var tzi = TimeZoneParser.ParseTimeZoneId(args[i]);
                                if (tzi == null)
                                {
                                    Console.WriteLine($"Invalid value for -changeTimezone '{args[i]}'. Use '-listTimezones' option to find valid values.");
                                    s_commandLineError = true;
                                    break;
                                }
                                photoFinisher.ChangeTimezoneTo = tzi;
                            }
                            break;

                        case "-updatefsdate":
                            photoFinisher.UpdateFileSystemDate = true;
                            break;

                        case "-log":
                            s_log = true;
                            break;

                        case "-listtimezones":
                        case "-listtimezone":
                            s_listTimezones = true;
                            break;

                        default:
                            Console.WriteLine($"Command-line syntax error: '{args[i]}' is not a recognized command.");
                            Console.WriteLine("Use '-h' for syntax help");
                            s_commandLineError = true;
                            break;
                    }
                }

                if (photoFinisher.AutoSort && string.IsNullOrEmpty(photoFinisher.DestinationDirectory))
                {
                    Console.WriteLine("Command-line error: '-sort' option requires '-d' destination option.");
                    Console.WriteLine("Use '-h' for syntax help");
                    s_commandLineError = true;
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

        static void PerformOperations(PhotoFinisher photoFinisher)
        {
            if (s_commandLineError) return;

            if (s_showSyntax)
            {
                Console.WriteLine(c_Syntax);
                return;
            }

            if (s_listTimezones)
            {
                TimeZoneParser.ListTimezoneIds();
                return;
            }

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