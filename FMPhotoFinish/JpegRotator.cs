using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Diagnostics;
using System.IO;

namespace FMPhotoFinish
{
    static class JpegRotator
    {
        const Int32 c_propId_Orientation = 0x0112;

        public static void RotateToVertical(string filename)
        {
            // Write the image to a temporary file in the same folder as the existing file
            string filenameTemp = filename + ".temp";

            // Load the image to rotate
            using (var image = Image.FromFile(filename))
            {
                // Get the existing orientation
                var piOrientation = image.GetPropertyItem(c_propId_Orientation);
                Debug.Assert(piOrientation.Id == c_propId_Orientation);
                Debug.Assert(piOrientation.Type == 3);
                Debug.Assert(piOrientation.Len == 2);

                // If it's aready vertical, do nothing
                if (piOrientation.Value[0] == 1) return;

                // Set the encoder value according to existing orientation
                EncoderValue ev;
                switch (piOrientation.Value[0])
                {
                    case 1: // Normal
                        return; // No rotation necessary, do nothing

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

                    default:
                        return; // It's in an orientation we don't know how to deal with such as transverse or transpose
                }

                // Change the orientation to 1 (normal) as we will rotate during the export
                piOrientation.Value[0] = 1;
                image.SetPropertyItem(piOrientation);

                // Prep the encoder and its parameters
                var encoder = System.Drawing.Imaging.Encoder.Transformation;
                var encParam = new EncoderParameter(encoder, (long)ev);
                var encParams = new EncoderParameters(1);
                encParams.Param[0] = encParam;

                // Write the image with a rotation transformation
                image.Save(filenameTemp, JpegCodecInfo, encParams);
            }

            // Delete the original file
            File.Delete(filename);

            // Rename the new one to the old name
            File.Move(filenameTemp, filename);
        }

        static ImageCodecInfo s_jpegCodecInfo = null;

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
