using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Media.Imaging;

namespace FMPhotoFinish
{
    /// <summary>
    /// Converts images to JPEG format
    /// </summary>
    class JpegConverter
    {

        /// <summary>
        /// Images from any format supported by Microsoft Windows Imaging Component (WIC) to JPEG.
        /// </summary>
        /// <remarks>
        /// .heic files are supported if the Windows HEIC and HEIF codecs are installed.
        /// Metadata is preserved.
        /// </remarks>
        public static void ConvertToJpeg(string srcFilename, string dstFilename)
        {
            // First, copy out all of the writable metadata properties
            var metadata = new Dictionary<Interop.PropertyKey, object>();
            using (var ps = WinShell.PropertyStore.Open(srcFilename))
            {
                int count = ps.Count;
                for (int i=0; i<count; ++i)
                {
                    var key = ps.GetAt(i);
                    if (ps.IsPropertyWriteable(key))
                    {
                        metadata.Add(key, ps.GetValue(key));
                    }
                }
            }

            // Set the rotation transformation appropriately
            var rotation = Rotation.Rotate0;
            bool flipHorizontal = false;
            bool flipVertical = false;
            if (metadata.ContainsKey(PropertyKeys.Orientation))
            {
                switch ((ushort)metadata[PropertyKeys.Orientation])
                {
                    case 2: // FlipHorizontal
                        flipHorizontal = true;
                        break;

                    case 3: // Rotated 180
                        rotation = Rotation.Rotate180;
                        break;

                    case 4: // FlipVertical
                        flipVertical = true;
                        break;

                    case 6: // Rotated 270
                        rotation = Rotation.Rotate90;
                        break;

                    case 8: // Rotated 90
                        rotation = Rotation.Rotate270;
                        break;
                }

                metadata[PropertyKeys.Orientation] = (ushort)1;
            }

            // Convert the image
            try
            {
                using (var stream = new FileStream(srcFilename, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    var decoder = BitmapDecoder.Create(stream, BitmapCreateOptions.None, BitmapCacheOption.Default);
                    BitmapFrame frame = decoder.Frames[0];

                    // If rotating, then strip any thumbnail
                    if (rotation != Rotation.Rotate0 || flipHorizontal || flipVertical)
                    {
                        frame = BitmapFrame.Create((BitmapSource)frame, null, (BitmapMetadata)frame.Metadata, frame.ColorContexts);
                    }

                    // Encode to JPEG
                    using (var dstStream = new FileStream(dstFilename, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                    {
                        var encoder = new JpegBitmapEncoder();
                        encoder.Rotation = rotation;
                        encoder.FlipHorizontal = flipHorizontal;
                        encoder.FlipVertical = flipVertical;
                        encoder.Frames.Add(frame);
                        encoder.Save(dstStream);
                    }
                }
            }
            catch (Exception err)
            {
                throw new ApplicationException($"Failed to convert image {srcFilename}.", err);
            }

            // Update the metadata
            using (var ps = WinShell.PropertyStore.Open(dstFilename, true))
            {
                foreach(var pair in metadata)
                {
                    if (ps.IsPropertyWriteable(pair.Key)) // Might have been writeable at the source format but not on the destination
                    {
                        ps.SetValue(pair.Key, pair.Value);
                    }
                }
                ps.Commit();
            }
        }
    }
}
