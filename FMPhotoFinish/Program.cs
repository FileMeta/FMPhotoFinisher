using System;
using System.IO;
using System.Text;
using FileMeta;

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
 *  -allTheWay
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

  -selectAfter <dateTime> Only select files from the source with a DateTaken
                   (for photos) or equivalent (for video and audio) after
                   the specified date/time. If the file does not have date
                   metadata then it will not be included. If timezone is not
                   specified then local timezone is assumed.

  -selectIncremental Select files with a DateTaken value that is newer than
                   the newest file found in the last incremental selection.
                   A file, FMPhotoFinisher_Incremental.json, will be written
                   in the destination to keep track of the date of the newest
                   selected file. When combined with ""-selectAfter"" then
                   the most recent date from the two options will be used.

Destination:
  If no destination is specified, then the updates are made in-place.

  -d <path>        A path to the folder where the processed files should be
                   placed.

  -sortby <patt>   Auto-sort the transferred files into a directory tree
                   rooted in the folder specified by -d. Supported patterns
                   are y (year), ym (year/month), ymd (year/month/day),
                   ymds (year/month/day subject).

  -sort            Equivalent to -sortby ymd

  -move            Move the files from the source to the destination. If this
                   parameter is not included then the unprocessed original
                   files are left at the source.

Standalone Operations:

  -listTimezones   List all timezone IDs and associated offsets.

  -authOneDrive <sourceName>
                   Interactively log in to a Microsoft OneDrive account,
                   authorize for access to the photos stored there, and
                   define a new named source on the local machine and
                   account with securely stored credentials.

  -listNamedSources
                   List named sources such as those created for OneDrive.

  -deleteNamedSource <sourceName>
                   Delete a pre-configured source such as one created by
                   onedrive

Operations:
  -autorot         Using the 'orientation' metadata flag, auto rotate images
                   to their vertical position and clear the orientation flag.

  -orderedNames    Canon cameras prefix photos with ""IMG_"" and videos with
                   ""MVI_"". This option renames videos to use the ""IMG_""
                   prefix thereby having them show in order with the
                   associated photos.

  -filenameFromMetadata
                   Names the files according to the date the media was
                   captured (e.g. photo was taken) plus subject and title
                   metadata. If date metadata is not present, does nothing
                   even if subject and title are present. See the Metadata
                   Bearing Filename Pattern below.

  -metadataFromFilename
                   Set subject, title, and keywords from filename contents.
                   Subject and title will only be set if there they don't
                   have existing values. Keywords will be added if matching
                   values don't already exist. See the Metadata Bearing
                   Filename Pattern below.

  -metadataFromFilenameOverwrite
                   Set subject, title, and keywords from filename contents.
                   New values will overwrite any existing subject and title
                   values. Keywords will be added if matching values don't
                   already exist. See teh Metadata Bearing Filename Pattern
                   below.

  -saveOriginalFn  Save the original filename in a custom metaTag stored
                   in the comments property. If the property already exists,
                   it is preserved.

  -setUUID         Set a unique UUID as a custom metaTag stored in the
                   comments property. If the property already exists, it is
                   preserved.

  -transcode       Transcode video and audio files to the preferred format
                   which is .mp4 for video and .m4a for audio. Also renames
                   .jpeg files to .jpg.

  -determineDate   Determines the date and timezone from available data and
                   sets the corresponding metadata fields. This operation is
                   automatically performed regardless of the command-line
                   flag if any other metadata fields are updated.

  -allTheWay       Performs all of the basic operations. Equivalent to:
                   -orderedNames -autorot -transcode -saveOriginalFn
                   -setUUID -determineDate

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

  -updateFsCreate  Update the file system dateCreated to match the metadata
                   Date Created / Date Taken value.

  -updateFsMod     Update the file system dateModified to match the metadata
                   Date Created / Date Modified value. Note that some cloud
                   sync systems key off the DateModified value and this may
                   interfere with that operation.

  -tag <tag>       Add a keyword tag to each file. This parameter may be
                   repeated to add multiple tags.

  -deduplicate     Remove duplicate media files. When copying, (using the
                   -d argument) duplicates are simply not copied.
                   If in-place or moving (using the -move argument)
                   duplicates are deleted.

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

Metadata Bearing Filename Pattern
  -metadataFromFilename, -metadataFromFilenameOverwrite, and
  -filenameFromMetadata use the following patterns for the filename.

  <date or DCF> <subject> - <title> #<tag> (<index>).<extension>

  Date or DCF is a date, number, or DCF name (Design Rule for Camera Name).
  Subject and title are separated by space-dash-space. A # hash sign
  precedes keywords (hashtags).

  An optional index number may appear at the
  end in parentheses. The number is used to prevent identical filenames
  when the metadata are the same.

  All components are optional. For example, the dash will only appear if
  there is a title. The hash sign will only appear if there are keywords.

  When reading metadata from the filename (-metadataFromFilename or 
  -metadataFromFilenameOverwrite) the date or DCF name prefix is ignored.
  However, -determineDate may use a date prefix in the process of
  finding the best source for when the photo was taken. There may be
  multiple hashtags following the # value.

  When setting the filename (-filenameFromMetadata), the format will be:

  yyyy-mm-dd_hhmmss <subject> - <title>.<extension>
  For example: 2019-01-15_142022 Mirror Lake - John Fishing.jpg

  If no title metadata is present, no dash or title will appear. If the
  subject is not present but a title is, then the date will be followed
  by the dash (surrounded by spaces) and the title.

  For example: 2019-01-15_142022 - John Fishing.jpg
";

        enum Operation
        {
            None,                   // Do nothing
            CommandLineError,       // A command-line error occurred. Do nothing.
            ShowHelp,               // Display help/syntax string
            ProcessMediaFiles,      // Default - process a set of media files
            ListTimeZones,          // List timezone names
            ListNamedSources,       // List named source on the local machine and account
            DeleteNamedSource,      // Delete a named source
            AuthOneDrive,           // Authorize a OneDrive named source
#if DEBUG
            TestAction,
#endif
        }

        const string c_logsFolder = "logs";
        const string c_logFileSuffix = "FMPhotoFinish Log.txt";

        static Operation s_operation;
        static SourceConfiguration s_sourceConfiguration;
        static IMediaSource s_mediaSource;
        static PhotoFinisher s_photoFinisher;
        static bool s_waitBeforeExit;
        static bool s_log;
        static TextWriter s_logWriter;
        static string s_sourceName;
#if DEBUG
        static Action<string> s_testAction;
        static string s_testArgument;
#endif

        static void Main(string[] args)
        {
            try
            {
                ParseCommandLine(args);

                switch (s_operation)
                {
                    case Operation.ShowHelp:
                        Console.WriteLine(c_Syntax);
                        break;

                    case Operation.ProcessMediaFiles:
                        ProcessMediaFiles();
                        break;

                    case Operation.ListTimeZones:
                        TimeZoneParser.ListTimezoneIds();
                        break;

                    case Operation.ListNamedSources:
                        break;

                    case Operation.DeleteNamedSource:
                        break;

                    case Operation.AuthOneDrive:
                        NamedSource.OneDriveLoginAndAuthorize(s_sourceName);
                        break;

#if DEBUG
                    case Operation.TestAction:
                        s_testAction(s_testArgument);
                        break;
#endif
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

        static void ParseCommandLine(string[] args)
        {
            s_sourceConfiguration = new SourceConfiguration();
            Program.s_photoFinisher = new PhotoFinisher();

            Program.s_photoFinisher.ProgressReported += ReportProgress;
            Program.s_photoFinisher.StatusReported += ReportStatus;


            try
            {
                if (args.Length == 0)
                {
                    s_operation = Operation.ShowHelp;
                    return;
                }

                for (int i = 0; i < args.Length; ++i)
                {
                    switch (args[i].ToLowerInvariant())
                    {
                        case "-h":
                        case "-?":
                            s_operation = Operation.ShowHelp;
                            return;

                        case "-s":
                            if (s_mediaSource != null) throw new Exception("CommandLine -s: Can only specify one source.");
                            s_mediaSource = new FileSource(NextArgument(args, ref i, "-s"), false);
                            s_operation = Operation.ProcessMediaFiles;
                            break;

                        case "-st":
                            if (s_mediaSource != null) throw new Exception("CommandLine -st: Can only specify one source.");
                            s_mediaSource = new FileSource(NextArgument(args, ref i, "-s"), true);
                            s_operation = Operation.ProcessMediaFiles;
                            break;

                        case "-sdcim":
                            if (s_mediaSource != null) throw new Exception("CommandLine -sdcim: Can only specify one source.");
                            s_mediaSource = new DcimSource();
                            s_operation = Operation.ProcessMediaFiles;
                            break;

#if DEBUG
                        case "-copydcim":
                            {
                                Console.WriteLine("Copying test files.");
                                int n = DcimSource.CopyDcimTestFiles();
                                Console.WriteLine($"Copied {n} files from DCIM_Test to DCIM.");
                            }
                            break;
#endif

                        case "-selectafter":
                            s_sourceConfiguration.SelectAfter = NextArgumentAsDate(args, ref i, "-selectAfter")
                                .ResolveTimeZone(TimeZoneInfo.Local).Date;
                            break;

                        case "-selectincremental":
                            s_sourceConfiguration.SelectIncremental = true;
                            break;

                        case "-d":
                            {
                                string dst = NextArgument(args, ref i, "-d");
                                if (!Directory.Exists(dst))
                                {
                                    throw new ArgumentException($"Destination folder '{dst}' does not exist.");
                                }
                                s_sourceConfiguration.DestinationDirectory = Path.GetFullPath(dst);
                                s_photoFinisher.DestinationDirectory = s_sourceConfiguration.DestinationDirectory;
                            }
                            break;

                        case "-w":
                            s_waitBeforeExit = true;
                            break;

                        case "-sortby":
                            switch (NextArgument(args, ref i, "-sortby").ToLowerInvariant())
                            {
                                case "y":
                                    s_photoFinisher.SortBy = DatePathType.Y;
                                    break;
                                case "ym":
                                    s_photoFinisher.SortBy = DatePathType.YM;
                                    break;
                                case "ymd":
                                    s_photoFinisher.SortBy = DatePathType.YMD;
                                    break;
                                case "ymds":
                                    s_photoFinisher.SortBy = DatePathType.YMDS;
                                    break;
                                default:
                                    throw new ArgumentException($"Unexpected value for -sortby: {args[i]}.");
                            }
                            break;

                        case "-sort":
                            s_photoFinisher.SortBy = DatePathType.YMD;
                            break;

                        case "-move":
                            s_sourceConfiguration.MoveFiles = true;
                            break;

                        case "-autorot":
                            s_photoFinisher.AutoRotate = true;
                            break;

                        case "-orderednames":
                            s_photoFinisher.SetOrderedNames = true;
                            s_photoFinisher.SetMetadataNames = false;
                            break;

                        case "-filenamefrommetadata":
                            s_photoFinisher.SetMetadataNames = true;
                            s_photoFinisher.SetOrderedNames = false;
                            break;

                        case "-metadatafromfilename":
                            s_photoFinisher.MetadataFromFilename = SetMode.SetIfEmpty;
                            break;

                        case "-metadatafromfilenameoverwrite":
                            s_photoFinisher.MetadataFromFilename = SetMode.SetAlways;
                            break;

                        case "-saveoriginalfn":
                        case "-saveoriginalfilename":
                            s_photoFinisher.SaveOriginalFilaname = true;
                            break;

                        case "-setuuid":
                            s_photoFinisher.SetUuid = true;
                            break;

                        case "-transcode":
                            s_photoFinisher.Transcode = true;
                            break;

                        case "-tag":
                            s_photoFinisher.AddKeywords.Add(NextArgument(args, ref i, "-tag"));
                            break;

                        case "-determinedate":
                            s_photoFinisher.AlwaysSetDate = true;
                            break;

                        case "-alltheway":
                            s_photoFinisher.AutoRotate = true;
                            if (!s_photoFinisher.SetMetadataNames)
                                s_photoFinisher.SetOrderedNames = true;
                            s_photoFinisher.SaveOriginalFilaname = true;
                            s_photoFinisher.SetUuid = true;
                            s_photoFinisher.Transcode = true;
                            s_photoFinisher.AlwaysSetDate = true;
                            break;

                        case "-deduplicate":
                            s_photoFinisher.DeDuplicate = true;
                            break;

                        case "-setdate":
                            s_photoFinisher.SetDateTo = NextArgumentAsDate(args, ref i, "-setdate")
                                .ResolveTimeZone(TimeZoneInfo.Local);
                            break;

                        case "-shiftdate":
                            {
                                ++i;
                                if (i >= args.Length)
                                {
                                    throw new ArgumentException($"Expected value for -shiftDate command-line argument.");
                                }

                                string timeShift = NextArgument(args, ref i, "-shiftdate");

                                // If the next argument starts with a sign, then the shift amount is a simple timespan.
                                if (timeShift[0] == '+' || timeShift[0] == '-')
                                {
                                    string s = (timeShift[0] == '+') ? timeShift.Substring(1) : timeShift;
                                    TimeSpan ts;
                                    if (!TimeSpan.TryParse(s, System.Globalization.CultureInfo.InvariantCulture, out ts))
                                    {
                                        throw new ArgumentException($"Invalid value for -shiftDate '{timeShift}'.");
                                    }

                                    s_photoFinisher.ShiftDateBy = ts;
                                }

                                // Else, shift amount is the difference of two times
                                else
                                {
                                    string secondDate = NextArgument(args, ref i, "-shiftdate second value");

                                    FileMeta.DateTag dtTarget;
                                    if (!FileMeta.DateTag.TryParse(timeShift, out dtTarget))
                                    {
                                        throw new ArgumentException($"Invalid value for -shiftDate '{args[i]}'.");
                                    }

                                    FileMeta.DateTag dtSource;
                                    if (!FileMeta.DateTag.TryParse(secondDate, out dtSource))
                                    {
                                        throw new ArgumentException($"Invalid value for -shiftDate '{args[i]}'.");
                                    }

                                    // Resolve timezone if it was ambiguous.
                                    dtTarget.ResolveTimeZone(TimeZoneInfo.Local);
                                    dtSource.ResolveTimeZone(TimeZoneInfo.Local);

                                    // For whatever reason, they might have used different timezones. Take the difference between the UTC versions.
                                    s_photoFinisher.ShiftDateBy = dtTarget.DateUtc.Subtract(dtSource.DateUtc);
                                }
                            }
                            break;

                        case "-settimezone":
                            {
                                string tz = NextArgument(args, ref i, "-settimezone");
                                var tzi = TimeZoneParser.ParseTimeZoneId(tz);
                                if (tzi == null)
                                {
                                    throw new ArgumentException($"Invalid value for -setTimezone '{tz}'. Use '-listTimezones' option to find valid values.");
                                }
                                s_photoFinisher.SetTimezoneTo = tzi;
                            }
                            break;

                        case "-changetimezone":
                            {
                                string tz = NextArgument(args, ref i, "-settimezone");
                                var tzi = TimeZoneParser.ParseTimeZoneId(tz);
                                if (tzi == null)
                                {
                                    throw new ArgumentException($"Invalid value for -changeTimezone '{tz}'. Use '-listTimezones' option to find valid values.");
                                }
                                s_photoFinisher.ChangeTimezoneTo = tzi;
                            }
                            break;

                        case "-updatefscreate":
                            s_photoFinisher.UpdateFileSystemDateCreated = true;
                            break;

                        case "-updatefsmod":
                            s_photoFinisher.UpdateFileSystemDateModified = true;
                            break;

                        case "-log":
                            s_log = true;
                            break;

                        case "-listtimezones":
                        case "-listtimezone":
                            s_operation = Operation.ListTimeZones;
                            break;

                        case "-authonedrive":
                            s_sourceName = NextArgument(args, ref i, "-authOneDrive");
                            s_operation = Operation.AuthOneDrive;
                            break;

#if DEBUG
                        case "-testaccess":
                            s_testAction = NamedSource.TestAccess;
                            s_testArgument = NextArgument(args, ref i, "-testaccess");
                            s_operation = Operation.TestAction;
                            break;

                        case "-testmetadatafromfilename":
                            s_testAction = MediaFile.TestMetadataFromFilename;
                            s_testArgument = NextArgument(args, ref i, "-testmetadatafromfilename");
                            s_operation = Operation.TestAction;
                            break;
#endif

                        default:
                            throw new ArgumentException($"Unexpected command-line argument: {args[i]}");
                    }
                }

                if (s_photoFinisher.SortBy != DatePathType.None && string.IsNullOrEmpty(s_photoFinisher.DestinationDirectory))
                {
                    throw new ArgumentException("'-sort' option requires '-d' destination option.");
                }
            }
            catch (Exception err)
            {
                Console.WriteLine(err.Message);
                Console.WriteLine("Use '-h' for syntax help");
                s_operation = Operation.CommandLineError;
            }

        } // ParseCommandLine

        static string NextArgument(string[] args, ref int i, string argName)
        {
            ++i;
            if (i >= args.Length)
            {
                throw new ArgumentException($"Expected value for {argName} command-line argument.");
            }
            return args[i];
        }

        static FileMeta.DateTag NextArgumentAsDate(string[] args, ref int i, string argName)
        {
            var date = NextArgument(args, ref i, argName);
            FileMeta.DateTag dt;
            if (!FileMeta.DateTag.TryParse(date, out dt))
            {
                throw new ArgumentException($"Invalid value for {argName} '{args[i]}'.");
            }
            return dt;
        }


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

        static void ProcessMediaFiles()
        {
            // Prepare logfile
            if (s_log)
            {
                string logDir = s_sourceConfiguration.DestinationDirectory;
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
            s_mediaSource.RetrieveMediaFiles(s_sourceConfiguration, s_photoFinisher);
            s_photoFinisher.PerformOperations();
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