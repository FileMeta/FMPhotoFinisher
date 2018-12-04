using System;
using WinShell;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;
using System.Text;

namespace FMPhotoFinish
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
        const string c_ffmpegVideoSettings = "-pix_fmt yuv420p -c:v libx264 -profile:v main -level:v 3.1 -crf 18";
        const string c_ffmpegAudioSettings = "-c:a aac"; // Let it use the default quality settings

        const string c_jpgExt = ".jpg";
        const string c_mp4Ext = ".mp4";
        const string c_m4aExt = ".m4a";

        #region Static Members

        static Encoding s_Utf8NoBOM = new UTF8Encoding(false);

        static Dictionary<string, MediaType> s_mediaExtensions = new Dictionary<string, MediaType>()
        {
            {c_jpgExt, MediaType.Image},
            {".jpeg", MediaType.Image},
            {c_mp4Ext, MediaType.Video},
            {".avi", MediaType.Video},
            {".mov", MediaType.Video},
            {".mpg", MediaType.Video},
            {".mpeg", MediaType.Video},
            {c_m4aExt, MediaType.Audio},
            {".mp3", MediaType.Audio},
            {".wav", MediaType.Audio}
        };

        static HashSet<string> s_isomExtensions = new HashSet<string>()
        {
            c_mp4Ext, ".mov", c_m4aExt
        };

        // Preferred formats (by media type
        static string[] s_preferredExtensions = new string[]
        {
            null,   // Unknown
            c_jpgExt, // Image
            c_mp4Ext, // Video
            c_m4aExt  // Audio
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

        /// <summary>
        /// Parses a timezone string into a signed integer representing the number of minutes offset from UTC.
        /// </summary>
        /// <param name="s">The timezone string to parse.</param>
        /// <param name="result">The parsed timezone.</param>
        /// <returns>True if successful, else false.</returns>
        /// <remarks>Example timezone values: "-05:00" "-5", "+06:00", "+6", "+10:30".</remarks>
        public static bool TryParseTimeZone(string s, out int result)
        {
            result = 0;
            var parts = s.Split(':');
            if (parts.Length < 1 || parts.Length > 2) return false;
            int n;
            if (!int.TryParse(parts[0], out n)) return false;
            result = n * 60;
            if (parts.Length > 1)
            {
                if (!int.TryParse(parts[1], out n)) return false;
                if (s[0] == '-') n = -n;
                result += n;
            }
            return true;
        }

        #endregion

        #region Delegates

        public delegate void ProgressReporter(string progress);

        #endregion


        string m_filepath;
        MediaType m_mediaType;

        // Writable Values from the Windows Property System
        // TODO: When all is done, are these used anyway?
        Dictionary<PROPERTYKEY, object> m_propsToSet = new Dictionary<PROPERTYKEY, object>();

        // Values that may come from any of several sources
        string m_make;
        string m_model;
        string m_imageUniqueId;
        int? m_timezone; // In minutes offset from UTC, positive or negative

        // Critical values from the Windows Property System
        TimeSpan? m_psDuration;

        // Values from the ISOM container (.MOV, .MP4, and .M4A formats)
        DateTime? m_isomCreationTime = null;
        DateTime? m_isomModificationTime = null;
        TimeSpan? m_isomDuration = null;

        public MediaFile(string filepath)
        {
            m_filepath = filepath;
            string ext = Path.GetExtension(filepath).ToLowerInvariant();
            if (!s_mediaExtensions.TryGetValue(ext, out m_mediaType))
            {
                throw new InvalidOperationException($"Media type '{ext}' is not supported.");
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

                    // System.Media.Duration
                    if (pk.Equals(PropertyKeys.Duration))
                    {
                        var duration = (ulong)propstore.GetValue(pk);
                        m_psDuration = new TimeSpan((long)duration);
                    }

                }
            }

            // Load Isom Properties
            // TODO: Verify that these are used.
            if (s_isomExtensions.Contains(ext))
            {
                var isom = FileMeta.IsomCoreMetadata.TryOpen(filepath);
                if (isom != null)
                {
                    using (isom)
                    {
                        m_isomCreationTime = isom.CreationTime;
                        m_isomModificationTime = isom.ModificationTime;
                        m_isomDuration = isom.Duration;
                    }
                }
            }

            // Load ExifTool Properties
            {
                var exifProperties = new List<KeyValuePair<string, string>>();
                s_exifTool.GetProperties(m_filepath, exifProperties);
                string software = null;

                foreach (var pair in exifProperties)
                {
                    int colon = pair.Key.IndexOf(':');
                    string key = (colon >= 0) ? pair.Key.Substring(colon + 1) : pair.Key;
                    switch (key.ToLowerInvariant())
                    {
                        case "timezone":
                            if (m_timezone == null)
                            {
                                int tz;
                                if (TryParseTimeZone(pair.Value, out tz))
                                {
                                    m_timezone = tz;
                                }
                            }
                            break;

                        case "make":
                            if (m_make == null)
                            {
                                m_make = pair.Value;
                            }
                            break;

                        case "model":
                            if (m_make == null)
                            {
                                m_model = pair.Value;
                            }
                            break;

                        // For AVI Video
                        case "software":
                            software = pair.Value;
                            break;
                    }
                }

                if (m_make == null) m_make = software;
                if (m_model == null) m_model = software;
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

        public bool IsPreferredFormat { get { return string.Equals(Path.GetExtension(m_filepath), PreferredFormat, StringComparison.OrdinalIgnoreCase); } }

        public bool TranscodeToPreferredFormat(ProgressReporter reporter)
        {
            string ext = Path.GetExtension(m_filepath);
            if (string.Equals(ext, PreferredFormat, StringComparison.OrdinalIgnoreCase))
                return true;

            if (m_mediaType == MediaType.Image)
            {
                Debug.Assert(ext.Equals(".jpeg", StringComparison.OrdinalIgnoreCase));
                ChangeExtensionTo(c_jpgExt);
                return true;
            }

            if (m_mediaType != MediaType.Audio && m_mediaType != MediaType.Video)
                return false;

            return Transcode(reporter);
        }

        #region IDisposable Support

        protected virtual void Dispose(bool disposing)
        {
            if (m_filepath == null) return;

            if (m_propsToSet.Count > 0)
            {
#if DEBUG
                Debug.WriteLine(Path.GetFileName(m_filepath));
                foreach (var pair in m_propsToSet)
                {
                    Debug.WriteLine($"   {pair.Value}");
                }
#endif
                using (var ps = PropertyStore.Open(m_filepath, true))
                {
                    foreach (var pair in m_propsToSet)
                    {
                        ps.SetValue(pair.Key, pair.Value);
                    }
                    ps.Commit();
                }
            }

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
            }
            catch
            {
                result = false;
            }
            return result;
        }

        const string c_FFMpeg = "FFMpeg.exe";
               
        bool Transcode(ProgressReporter reporter)
        {
            // If inbound file does not have a duration, it's not a transcodeable media file
            if (m_psDuration == null)
            {
                return false;
            }

            string newPath = Path.ChangeExtension(m_filepath, PreferredFormat);
            MakeFilepathUnique(ref newPath);

            Process transcoder = null;
            bool result = false;
            try
            {
                // Compose arguments
                string arguments;
                if (m_mediaType == MediaType.Video)
                {
                    arguments = $"-hide_banner -i {m_filepath} {c_ffmpegVideoSettings} {c_ffmpegAudioSettings} {newPath}";
                }
                else if (m_mediaType == MediaType.Audio)
                {
                    arguments = $"-hide_banner -i {m_filepath} {c_ffmpegAudioSettings} {newPath}";
                }
                else
                {
                    throw new InvalidOperationException();
                }

                // Prepare process start
                var psi = new ProcessStartInfo(c_FFMpeg, arguments);
                psi.UseShellExecute = false;
                psi.CreateNoWindow = true; // Set to false if you want to monitor
                psi.RedirectStandardError = true;
                psi.StandardErrorEncoding = s_Utf8NoBOM;

                transcoder = Process.Start(psi);

                bool wroteProgress = false;
                var sb = new StringBuilder();
                for (; ; )
                {
                    int i = transcoder.StandardError.Read();
                    if (i < 0) break;
                    if (i == '\n')
                    {
                        sb.Clear();
                    }
                    else if (i == '\r' && transcoder.StandardError.Peek() != '\n')
                    {
                        reporter?.Invoke("Transcoding: " + sb.ToString());
                        wroteProgress = true;
                        sb.Clear();
                    }
                    else
                    {
                        sb.Append((char)i);
                    }
                }
                transcoder.WaitForExit();
                if (wroteProgress)
                {
                    reporter?.Invoke(null);
                }

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
                    using (isom)
                    {
                        Debug.Assert(m_psDuration != null); // Should have exited early if duration is null.
                        if (isom.Duration == null
                            || Math.Abs(m_psDuration.Value.Ticks - isom.Duration.Ticks) > (250L * 10000L)) // 1/4 second
                        {
                            result = false;
                        }
                    }
                }
            }

            // If successful, replace original with transcoded. If failed, delete the transcoded version.
            if (result)
            {
                File.Delete(m_filepath);
                m_filepath = newPath;
                if (m_make != null)
                    m_propsToSet.Add(PropertyKeys.Make, m_make);
                if (m_model != null)
                    m_propsToSet.Add(PropertyKeys.Model, m_model);
            }
            else
            {
                File.Delete(newPath);
            }

            return result;
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

#region Init and Shutdown

        static PropertySystem s_propSystem = new PropertySystem();
        static ExifToolWrapper.ExifTool s_exifTool = new ExifToolWrapper.ExifTool();
        static readonly StaticDisposer s_psDisposer = new StaticDisposer();

        private sealed class StaticDisposer
        {
            ~StaticDisposer()
            {
                if (s_propSystem != null)
                {
                    s_propSystem.Dispose();
                    s_propSystem = null;
                }
                if (s_exifTool != null)
                {
                    s_exifTool.Dispose();
                    s_exifTool = null;
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
        public static PROPERTYKEY Duration = new PROPERTYKEY("64440490-4C8B-11D1-8B70-080036B11A03", 3);
        public static PROPERTYKEY Make = new PROPERTYKEY("14b81da1-0135-4d31-96d9-6cbfc9671a99", 271);
        public static PROPERTYKEY Model = new PROPERTYKEY("14b81da1-0135-4d31-96d9-6cbfc9671a99", 272);
    }

}
