﻿using System;
using WinShell;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;
using System.Text;

namespace FMPhotoFinisher
{
    enum MediaType
    {
        Unsupported = 0, // Unsupported type
        Image = 1,  // .jpg .jpeg
        Video = 2,  // .mp4 .avi .mov .mpg .mpeg
        Audio = 3   // .m4a .mp3 .wav 
    }

    /// <summary>
    /// Performs operations on a media file such as metadata changes, rotation, recoding, etc.
    /// </summary>
    class MediaFile : IDisposable
    {
        #region Static Members

        static Encoding s_Utf8NoBOM = new UTF8Encoding(false);

        static Dictionary<string, MediaType> s_mediaExtensions = new Dictionary<string, MediaType>()
        {
            {".jpg", MediaType.Image},
            {".jpeg", MediaType.Image},
            {".mp4", MediaType.Video},
            {".avi", MediaType.Video},
            {".mov", MediaType.Video},
            {".mpg", MediaType.Video},
            {".mpeg", MediaType.Video},
            {".m4a", MediaType.Audio},
            {".mp3", MediaType.Audio},
            {".wav", MediaType.Audio}
        };

        static HashSet<string> s_isomExtensions = new HashSet<string>()
        {
            ".mp4", ".mov", ".m4a"
        };

        // Preferred formats (by media type
        static string[] s_preferredExtensions = new string[]
        {
            null,   // Unknown
            ".jpg", // Image
            ".mp4", // Video
            ".m4a"  // Audio
        };

        public static MediaType GetMediaType(string filenameOrExtension)
        {
            int ext = filenameOrExtension.LastIndexOf('.');
            if (ext >= 0)
            {
                MediaType t;
                if (s_mediaExtensions.TryGetValue(filenameOrExtension.Substring(ext).ToLowerInvariant(), out t))
                {
                    return t;
                }
            }
            return MediaType.Unsupported;
        }

        public static bool IsSupportedMediaType(string filenameOrExtension)
        {
            return GetMediaType(filenameOrExtension) != MediaType.Unsupported;
        }

        public static void MakeFilepathUnique(ref string filepath)
        {
            if (!File.Exists(filepath)) return;

            string basepath = Path.Combine(Path.GetDirectoryName(filepath), Path.GetFileNameWithoutExtension(filepath));
            string extension = Path.GetExtension(filepath);

            int index = 1;

            // Strip and parse any parenthesized index
            int p = basepath.Length;
            if (p > 2 && basepath[p-1] == ')' && char.IsDigit(basepath[p-2]))
            {
                p -= 2;
                while (char.IsDigit(basepath[p - 1])) --p;
                if (p > 1 && basepath[p-1] == '(')
                {
                    --p;
                    // Found a parenthesized index
                    index = int.Parse(basepath.Substring(p + 1, basepath.Length - p - 2)) + 1; // Update the index to the next value
                    basepath = basepath.Substring(0, p); // Remove the parenthesized index from the basepath;
                }
            }

            // Generate a unique filename
            do
            {
                filepath = $"{basepath}({index}){extension}";
                ++index;
            } while (File.Exists(filepath));
        }

        #endregion

        #region Delegates

        public delegate void ProgressReporter(string progress);

        #endregion


        string m_filepath;
        string m_ext;
        MediaType m_mediaType;

        // Values from the Windows Property System
        Dictionary<PROPERTYKEY, object> m_properties = new Dictionary<PROPERTYKEY, object>();

        // Values from the ISOM container (.MOV, .MP4, and .M4A formats)
        DateTime? m_IsomCreationTime = null;
        DateTime? m_IsomModificationTime = null;
        TimeSpan? m_isomDuration = null;

        public MediaFile(string filepath)
        {
            m_filepath = filepath;
            m_ext = Path.GetExtension(filepath).ToLowerInvariant();
            if (!s_mediaExtensions.TryGetValue(m_ext, out m_mediaType))
            {
                throw new InvalidOperationException($"Media type '{m_ext}' is not supported.");
            }

            Orientation = 1; // Defaults to normal/vertical

            // Load Windows Property Store properties
            using (var propstore = PropertyStore.Open(filepath, false))
            {
                int nProps = propstore.Count;
                for (int i=0; i<nProps; ++i)
                {
                    var pk = propstore.GetAt(i);
                    if (pk.Equals(PropertyKeys.Orientation))
                    {
                        Orientation = (int)(ushort)propstore.GetValue(pk);
                    }
                    else if (IsCopyable(pk))
                    {
                        m_properties[pk] = propstore.GetValue(pk);
                    }
                }
            }

            // Load Isom Properties
            if (s_isomExtensions.Contains(m_ext))
            {
                var isom = FileMeta.IsomCoreMetadata.TryOpen(filepath);
                if (isom != null)
                {
                    using (isom)
                    {
                        m_IsomCreationTime = isom.CreationTime;
                        m_IsomModificationTime = isom.ModificationTime;
                        m_isomDuration = isom.Duration;
                    }
                }
            }
        }

        public string Filepath { get { return m_filepath; } }

        public MediaType MediaType { get { return m_mediaType; } }

        public int Orientation { get; private set; }
        
        public void RotateToVertical()
        {
            JpegRotator.RotateToVertical(m_filepath);
            Orientation = 1; // Normal
        }

        public string PreferredFormat { get { return s_preferredExtensions[(int)m_mediaType]; } }

        public bool IsPreferredFormat { get { return string.Equals(m_ext, PreferredFormat, StringComparison.Ordinal); } }

        public bool TranscodeToPreferredFormat(ProgressReporter reporter)
        {
            switch (m_ext)
            {
                case ".jpeg":
                    return ChangeExtensionTo(".jpg");

                case ".avi":
                case ".mov":
                case ".mpg":
                case ".mpeg":
                    return TranscodeVideo(reporter);

                case ".mp3":
                case ".wav":
                    return TranscodeAudio();
                
                // For all others do nothing
            }

            return true;
        }

        #region IDisposable Support

        protected virtual void Dispose(bool disposing)
        {
            if (m_filepath == null) return;

            if (!disposing)
            {
                System.Diagnostics.Debug.Fail("Failed to dispose of MediaFile.");
            }
        }

#if DEBUG
        ~MediaFile()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(false);
        }
#endif

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
#if DEBUG
            GC.SuppressFinalize(this);
#endif
        }

        #endregion // IDisposable Support

        #region Private Members

        bool ChangeExtensionTo(string newExt)
        {
            bool result = true;
            try
            {
                string newPath = Path.ChangeExtension(m_filepath, newExt);
                MakeFilepathUnique(ref newPath);
                File.Move(m_filepath, newPath);
                m_filepath = newPath;
                m_ext = newExt;
            }
            catch
            {
                result = false;
            }
            return result;
        }

        const string c_FFMpeg = "FFMpeg.exe";
        const string c_mp4Extension = ".mp4";
               
        bool TranscodeVideo(ProgressReporter reporter)
        {
            string newPath = Path.ChangeExtension(m_filepath, c_mp4Extension);
            MakeFilepathUnique(ref newPath);

            Process transcoder = null;
            bool result = false;
            try
            {
                // Compose arguments
                var arguments = $"-hide_banner -i {m_filepath} -pix_fmt yuv420p -c:v libx264 -profile:v main -level:v 3.1 -crf 18 -c:a aac -f mp4 {newPath}";

                // Prepare process start
                var psi = new ProcessStartInfo(c_FFMpeg, arguments);
                psi.UseShellExecute = false;
                psi.CreateNoWindow = true; // Set to false if you want to monitor
                psi.RedirectStandardError = true;
                psi.StandardErrorEncoding = s_Utf8NoBOM;

                transcoder = Process.Start(psi);

                for (; ; )
                {
                    var line = transcoder.StandardError.ReadLine();
                    if (line == null) break;
                    reporter?.Invoke("Transcoding: " + line);
                }
                transcoder.WaitForExit();

                result = transcoder.ExitCode == 0;
            }
            finally
            {
                if (transcoder != null)
                {
                    transcoder.Dispose();
                    transcoder = null;
                }
            }

            if (result)
            {
                // Confirm the transcoding by reading out the duration and comparing with original (if available)
                var isom = FileMeta.IsomCoreMetadata.TryOpen(newPath);
                if (isom == null)
                {
                    result = false;
                }
                else
                {
                    // TODO: Compare against duration data from other sources than just isom.

                    using (isom)
                    {
                        if (m_isomDuration != null)
                        {
                            if (Math.Abs(m_isomDuration.Value.Ticks - isom.Duration.Ticks) > (250L * 10000L)) // 1/4 second
                            {
                                result = false;
                            }
                        }
                    }
                }
            }

            // If successful, replace original with transcoded. If failed, delete the transcoded version.
            if (result)
            {
                File.Delete(m_filepath);
                m_filepath = newPath;
                m_ext = c_mp4Extension;
            }
            else
            {
                File.Delete(newPath);
            }

            return result;
        }

        bool TranscodeAudio()
        {
            return false;
        }

        #endregion // Private Members

        #region Private Static Members

        // Cache whether a propery is copyable
        static Dictionary<PROPERTYKEY, bool> s_propertyIsCopyable = new Dictionary<PROPERTYKEY, bool>();

        static bool IsCopyable(PROPERTYKEY pk)
        {
            bool result;
            if (s_propertyIsCopyable.TryGetValue(pk, out result))
            {
                return result;
            }

            var desc = s_propSystem.GetPropertyDescription(pk);
            result = desc != null
                && desc.ValueTypeIsSupported
                && (desc.TypeFlags & PROPDESC_TYPE_FLAGS.PDTF_ISINNATE) == 0;
            s_propertyIsCopyable[pk] = result;

            return result;
        }

        #endregion

        #region PropertyStore

        static PropertySystem s_propSystem = new PropertySystem();
        static readonly PropSysStaticDisposer s_psDisposer = new PropSysStaticDisposer();

        private sealed class PropSysStaticDisposer
        {
            ~PropSysStaticDisposer()
            {
                if (s_propSystem != null)
                {
                    s_propSystem.Dispose();
                    s_propSystem = null;
                }
            }
        }

        #endregion
    }

    /// <summary>
    /// Properties defined here: https://msdn.microsoft.com/en-us/library/windows/desktop/dd561977(v=vs.85).aspx
    /// </summary>
    static class PropertyKeys
    {
        public static PROPERTYKEY Orientation = new PROPERTYKEY("14B81DA1-0135-4D31-96D9-6CBFC9671A99", 274);
        public static PROPERTYKEY DateTaken = new PROPERTYKEY("14B81DA1-0135-4D31-96D9-6CBFC9671A99", 36867);

    }

}
