using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FMPhotoFinish
{
    interface IMediaQueue
    {
        /// <summary>
        /// Reports curent status of the operation (e.g. "3 of 5 copied").
        /// </summary>
        /// <remarks>
        /// This would typically be shown in a status bar. When sent to the console, it should be sent
        /// to stderr with a \r (not \r\n). A null or blank message clears the status.
        /// </remarks>
        void ReportStatus(string message);

        /// <summary>
        /// Reports messages about operations starting or completing - for the user to read.
        /// Typically shown in a log format.
        /// </summary>
        /// <param name="message">The message to report.</param>
        void ReportProgress(string message);

        /// <summary>
        /// Add a file to the media processing queue.
        /// </summary>
        /// <param name="pfi">Informaiton about the media file to be processed.</param>
        /// <remarks>During processing, the location of the file may change. In which case
        /// the <see cref="ProcessFileInfo.Filepath"/> property will be updated.</remarks>
        void Add(ProcessFileInfo pfi);

        /// <summary>
        /// Indicate to the media processor that this is a limited batch.
        /// </summary>
        /// <remarks>
        /// Some sources work better with limited batch sizes. Setting this value to true
        /// indicates that the source has more media files beyond the batch that it just
        /// processed and that the system should call again once this batch is complete.
        /// </remarks>
        bool RequestAnotherBatch { get; set; }
    }

    interface IMediaSource
    {
        /// <summary>
        /// Select the files to be processed and copy them if needed.
        /// </summary>
        /// <param name="mediaQueue">An instance of <see cref="IMediaQueue"/>.</param>
        /// <remarks>
        /// <para>Files should be selected, copied or moved to the destination directory specified
        /// by <see cref="IMediaQueue.DestinationFolder"/>, and added to the queue using
        /// <see cref="IMediaQueue.Add"/>.
        /// </para>
        /// <para>Files are selected from the source according to the criteria specified in
        /// the source-specific configuration.
        /// </para>
        /// <para>If <see cref="IMediaQueue.DestinationFolder"/> is not null, then files should be
        /// moved or copied into that folder. The choice of moving or copying is according to the
        /// source configuration and/or source-specific configuration.
        /// </para>
        /// <para>If IMediaQueue.DestinationFolder is null, then files are to be processed in-place
        /// and not moved. For some sources, that is not appropriate in which case an
        /// <see cref="InvalidOperationException"/> should be thrown.
        /// </para>
        /// <para>Each file is added to the queue through a call to <see cref="IMediaQueue.Add"/>.
        /// When processing is in-place, then filename and originalFilename will be the same.
        /// </para>
        /// <para>Progress should be reported using <see cref="IMediaQueue.ReportStatus"/>
        /// and <see cref="IMediaQueue.ReportProgress"/>.</para>
        /// </remarks>
        void RetrieveMediaFiles(SourceConfiguration sourceConfig, IMediaQueue mediaQueue);
    }
}
