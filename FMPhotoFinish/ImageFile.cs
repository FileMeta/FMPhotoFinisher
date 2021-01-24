using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace FMPhotoFinish
{
    static class ImageFile
    {
        const Int32 c_propId_Orientation = 0x0112;
        const EncoderValue c_encoderValueZero = (EncoderValue)0; // Actually the same as ColorTypeCMYK but that value is not used.

        /// <summary>
        /// Replace the file with a resized and righted version
        /// </summary>
        /// <param name="filename">The filename to update</param>
        /// <param name="width">The width of the resulting image in pixels or zero.</param>
        /// <param name="height">The height of the resulting image in pixels or zero.</param>
        /// <remarks>
        /// <para>If both width and height are zero, the size (resolution) will be unchanged though a rotation
        /// may result in the x and y dimensions being swapped.
        /// </para>
        /// <para>If either width or height is zero, the other dimension will be calculated so as to preserve
        /// the original aspect ratio, with rotation taken into account.</para>
        /// </remarks>
        public static void ResizeAndRightImage(String filename, int width, int height)
        {
            using (var stream = new FileStream(filename, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            {
                ResizeAndRightImage(stream, stream, width, height);
            }
        }

        /// <summary>
        /// Replace the file with a righted (turned vertical) version
        /// </summary>
        /// <param name="filename">The filename to update</param>
        public static void RightImage(String filename)
        {
            using (var stream = new FileStream(filename, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            {
                RightImage(stream, stream);
            }
        }

        /// <summary>
        /// Write a resized and righted image to the target stream.
        /// </summary>
        /// <param name="src">Source stream from which to read the image.</param>
        /// <param name="dst">Destination stream to which the image should be written.</param>
        /// <param name="width">The width of the resulting image in pixels or zero.</param>
        /// <param name="height">The height of the resulting image in pixels or zero.</param>
        /// <remarks>
        /// <para>Source and destionation streams may be the same in which case the file will be
        /// updated in-place.
        /// </para>
        /// <para>If both width and height are zero, the size (resolution) will be unchanged though a rotation
        /// may result in the x and y dimensions being swapped.
        /// </para>
        /// <para>If either width or height is zero, the other dimension will be calculated so as to preserve
        /// the original aspect ratio, with rotation taken into account.</para>
        /// </remarks>
        public static void ResizeAndRightImage(Stream src, Stream dst, int width, int height)
        {
            // Load the image to change
            using (var image = Image.FromStream(src))
            {
                if (width == 0 && height == 0)
                {
                    RightImage(src, dst);
                    return;
                }

                // targetWidth and targetHeight are the resized width and height
                // before any rotation.
                int targetWidth = width;
                int targetHeight = height;
                RotateFlipType rft = RotateFlipType.RotateNoneFlipNone;

                var imgprops = image.PropertyItems;

                // Check the orientation and determine whether image must be rotated
                {
                    // Image.GetPropertyItem throws an exception if the image is not Exif
                    // So we use the property collection instead
                    var orientation =
                    (from ip in imgprops
                     where ip.Id == c_propId_Orientation
                     select ip.Value).FirstOrDefault();

                    if (orientation != null)
                    {
                        switch (orientation[0])
                        {
                            // case 1: // Vertical
                                // do nothing
                                // break;
                            case 2: // FlipHorizontal
                                rft = RotateFlipType.RotateNoneFlipX;
                                break;
                            case 3: // Rotated 180
                                rft = RotateFlipType.Rotate180FlipNone;
                                break;
                            case 4: // FlipVertical
                                rft = RotateFlipType.Rotate180FlipX;
                                break;
                            case 5:
                                targetWidth = height;
                                targetHeight = width;
                                rft = RotateFlipType.Rotate90FlipX;
                                break;
                            case 6: // Rotated 270
                                targetWidth = height;
                                targetHeight = width;
                                rft = RotateFlipType.Rotate90FlipNone;
                                break;
                            case 8: // Rotated 90
                                targetWidth = height;
                                targetHeight = width;
                                rft = RotateFlipType.Rotate270FlipNone;
                                break;
                        }
                    }
                }

                Debug.Assert(targetWidth != 0 || targetHeight != 0);
                if (targetWidth == 0)
                    targetWidth = (int)((long)image.Width * (long)targetHeight / (long)image.Height);
                if (targetHeight == 0)
                    targetHeight = (int)((long)image.Height * (long)targetWidth / (long)image.Width);

                using (var resizedImage = new Bitmap(targetWidth, targetHeight))
                {
                    resizedImage.SetResolution(72, 72);

                    // Copy metadata and fix rotation.
                    foreach (var prop in imgprops)
                    {
                        if (prop.Id == c_propId_Orientation)
                        {
                            prop.Value[0] = 1;  // Set it back to vertical.
                        }
                        resizedImage.SetPropertyItem(prop);
                    }

                    using (var graphic = Graphics.FromImage(resizedImage))
                    {
                        graphic.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        graphic.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                        graphic.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                        graphic.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                        graphic.DrawImage(image, 0, 0, targetWidth, targetHeight);
                    }

                    if (rft != RotateFlipType.RotateNoneFlipNone)
                        resizedImage.RotateFlip(rft);

                    if (dst.CanSeek) dst.Position = 0;
                    resizedImage.Save(dst, image.RawFormat);
                    if (dst.CanSeek) dst.SetLength(dst.Position);
                }
            }
        }

        public static void RightImage(Stream src, Stream dst)
        {
            // Load the image to rotate
            using (var image = Image.FromStream(src))
            {
                // Get the existing orientation
                var piOrientation = image.GetPropertyItem(c_propId_Orientation);
                Debug.Assert(piOrientation.Id == c_propId_Orientation);
                Debug.Assert(piOrientation.Type == 3);
                Debug.Assert(piOrientation.Len == 2);

                // Set the encoder value according to existing orientation
                EncoderValue ev;
                switch (piOrientation.Value[0])
                {
                    case 2: // FlipHorizontal
                        ev = EncoderValue.TransformFlipHorizontal;
                        break;

                    case 3: // Rotated 180
                        ev = EncoderValue.TransformRotate180;
                        break;

                    case 4: // FlipVertical
                        ev = EncoderValue.TransformFlipVertical;
                        break;

                    case 6: // Rotated 270
                        ev = EncoderValue.TransformRotate90;
                        break;

                    case 8: // Rotated 90
                        ev = EncoderValue.TransformRotate270;
                        break;

                    case 1: // Normal
                    default:
                        src.CopyTo(dst);
                        return;
                }

                // Change the orientation to 1 (normal) as we will rotate during the export
                piOrientation.Value[0] = 1;
                image.SetPropertyItem(piOrientation);

                // Prep the encoder parameters
                var encParams = new EncoderParameters(1);
                encParams.Param[0] = new EncoderParameter(Encoder.Transformation, (long)ev);

                // Write the image with a rotation transformation
                image.Save(dst, JpegCodecInfo, encParams);
            }
        }

        static ImageCodecInfo s_jpegCodecInfo;

        static ImageCodecInfo JpegCodecInfo
        {
            get
            {
                if (s_jpegCodecInfo == null)
                {
                    foreach (var encoder in ImageCodecInfo.GetImageEncoders())
                    {
                        if (encoder.MimeType.Equals("image/jpeg", StringComparison.OrdinalIgnoreCase))
                        {
                            s_jpegCodecInfo = encoder;
                            break;
                        }
                    }
                    if (s_jpegCodecInfo == null)
                    {
                        throw new ApplicationException("Unable to locate GDI+ JPEG Encoder");
                    }
                }
                return s_jpegCodecInfo;
            }
        }

    }

}
