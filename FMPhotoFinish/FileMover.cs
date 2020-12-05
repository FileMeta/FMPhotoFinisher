using System;
using System.Collections.Generic;
using System.IO;

namespace FMPhotoFinish
{
    /// <summary>
    /// Helper class for file-based sources
    /// </summary>
    static class FileMover
    {
        public static void CopyOrMoveFiles(List<ProcessFileInfo> queue, SourceConfiguration sourceConfig, IMediaQueue mediaQueue)
        {
            string verb = sourceConfig.MoveFiles ? "Moving" : "Copying";
            mediaQueue.ReportProgress($"{verb} media files to working folder: {sourceConfig.DestinationDirectory}.");

            // Sum up the size of the files to be copied
            long selectedFilesSize = 0;
            foreach (var pfi in queue)
            {
                selectedFilesSize += pfi.Size;
            }

            uint startTicks = (uint)Environment.TickCount;
            long bytesCopied = 0;

            int n = 0;
            foreach (var pfi in queue)
            {
                if (bytesCopied == 0)
                {
                    mediaQueue.ReportStatus($"{verb} file {n + 1} of {queue.Count}");
                }
                else
                {
                    uint ticksElapsed;
                    unchecked
                    {
                        ticksElapsed = (uint)Environment.TickCount - startTicks;
                    }

                    double bps = ((double)bytesCopied * 1000.0) / (double)ticksElapsed;
                    double remaining = (selectedFilesSize - bytesCopied) / bps;
                    TimeSpan remain = new TimeSpan(((long)((selectedFilesSize - bytesCopied) / bps)) * 10000000L);

                    mediaQueue.ReportStatus($"{verb} file {n + 1} of {queue.Count}. Time remaining: {remain.FmtCustom()} MBps: {(bps / (1024 * 1024)):#,###.###}");
                }

                string dstFilepath = Path.Combine(sourceConfig.DestinationDirectory, Path.GetFileName(pfi.OriginalFilepath));
                MediaFile.MakeFilepathUnique(ref dstFilepath);

                if (sourceConfig.MoveFiles)
                {
                    File.Move(pfi.Filepath, dstFilepath);
                }
                else
                {
                    File.Copy(pfi.Filepath, dstFilepath);
                }
                pfi.Filepath = dstFilepath;
                bytesCopied += pfi.Size;
                ++n;

                // Add to the destination queue
                mediaQueue.Add(pfi);
            }

            TimeSpan elapsed;
            unchecked
            {
                uint ticksElapsed = (uint)Environment.TickCount - startTicks;
                elapsed = new TimeSpan(ticksElapsed * 10000L);
            }

            mediaQueue.ReportStatus(null);
            mediaQueue.ReportProgress($"{verb} complete. {queue.Count} files, {bytesCopied / (1024.0 * 1024.0): #,##0.0} MB, {elapsed.FmtCustom()} elapsed");
        }

        public static void EnqueueFiles(List<ProcessFileInfo> queue, IMediaQueue mediaQueue)
        {
            foreach (var pfi in queue)
            {
                mediaQueue.Add(pfi);
            }
        }

    }
}
