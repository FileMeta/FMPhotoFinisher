using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FMPhotoFinish
{
    class SourceConfiguration
    {
        /// <summary>
        /// A path to the destination folder where files should be placed. Null if a destination
        /// folder was not specified and the files will be processed in-place.
        /// </summary>
        public string DestinationDirectory { get; set; }

        /// <summary>
        /// True if the <see cref="IMediaSource"/> shuld move the files to the destination
        /// (as opposed to copying them).
        /// </summary>
        /// <remarks>For some implementations of IMediaSource a move is not an option
        /// and this property should be ignored.</remarks>
        public bool MoveFiles { get; set; }

        /// <summary>
        /// True if only files that weren't selected in previous runs should be selected
        /// in this run.
        /// </summary>
        public bool SelectIncremental { get; set; }

        /// <summary>
        /// If present, only select files created after the specified DateTime.
        /// </summary>
        public DateTime? SelectAfter { get; set; }

        /// <summary>
        /// Determine the date after which items should be selected based on bookmark and source path.
        /// </summary>
        /// <param name="sourcePath">The source path which is used to look up a bookmark.</param>
        /// <returns>A <see cref="DateTime?"/> which, if it has a value, is the date-time after which items should be selected.</returns>
        /// <remarks>
        /// Picks the latter of <see cref="SelectAfter"/> or the bookmark associated with
        /// sourcePath contingent on whether SelectAfter and <see cref="SelectIncremental"/>
        /// is set.
        /// </remarks>
        public DateTime? GetBookmarkOrAfter(string sourcePath)
        {
            // Determine the "after" threshold from SelectAfter and SelectIncremental
            DateTime? after = SelectAfter;
            if (SelectIncremental)
            {
                var bookmark = new IncrementalBookmark(DestinationDirectory);
                var incrementalAfter = bookmark.GetBookmark(sourcePath);
                if (incrementalAfter.HasValue)
                {
                    if (!after.HasValue || after.Value < incrementalAfter.Value)
                    {
                        after = incrementalAfter;
                    }
                }
            }
            return after;
        }

        /// <summary>
        /// Conditionally sets a bookmark for use in a future run.
        /// </summary>
        /// <param name="sourcePath">The source path associated with the bookmark.</param>
        /// <param name="latestFound">The dateTime of the latest item found.</param>
        /// <remarks>Only sets a bookmark if <see cref="SelectIncremental"/> is set AND
        /// sourcePath has a value. Otherwise, does nothing.
        /// </remarks>
        public bool SetBookmark(string sourcePath, DateTime? latestFound)
        {
            if (!SelectIncremental) return false;
            if (!latestFound.HasValue) return false;
            var bookmark = new IncrementalBookmark(DestinationDirectory);
            bookmark.SetBookmark(sourcePath, latestFound.Value);
            return true;
        }

    }

}
