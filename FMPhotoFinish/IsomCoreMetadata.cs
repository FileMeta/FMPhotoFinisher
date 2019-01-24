/*
---
name: IsomCoreMetadata.cs
description: Read and write basic metadata on ISO Base Media File Format including .MP4 .MOV and .M4A
url: https://raw.githubusercontent.com/FileMeta/IsomCoreMetadata/master/IsomCoreMetadata.cs
version: 1.2
keywords: CodeBit
dateModified: 2018-05-02
license: http://unlicense.org
# Metadata in MicroYaml format. See http://filemeta.org/CodeBit.html
...
*/

/*
Unlicense: http://unlicense.org

This is free and unencumbered software released into the public domain.

Anyone is free to copy, modify, publish, use, compile, sell, or distribute
this software, either in source code form or as a compiled binary, for any
purpose, commercial or non-commercial, and by any means.

In jurisdictions that recognize copyright laws, the author or authors of this
software dedicate any and all copyright interest in the software to the
public domain. We make this dedication for the benefit of the public at large
and to the detriment of our heirs and successors. We intend this dedication
to be an overt act of relinquishment in perpetuity of all present and future
rights to this software under copyright law.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN
ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

For more information, please refer to <http://unlicense.org/>
*/

/*
References:
   ISO Base Media File Format Specification: http://standards.iso.org/ittf/PubliclyAvailableStandards/c068960_ISO_IEC_14496-12_2015.zip
*/

using System;
using System.IO;
using System.Text;
using System.Diagnostics;
using System.Collections.Generic;

namespace FileMeta
{
    /// <summary>
    /// Class for retrieving and setting 'core' metadata on ISO Base Media File Format (ISOM) files including MP4, MOV, and M4A formats among others.
    /// </summary>
    /// <remarks>
    /// "Core Metadata" in this case referrs to metadata that are inherent in the data structures. These include
    /// the Brand, Version, CompatibleBrands, CreationTime, ModifiedTime, and Duration. Updates to CreationTime and ModifiedTime
    /// are supported.
    /// </remarks>
    public class IsomCoreMetadata : IDisposable
    {
        Stream m_stream;
        bool m_leaveOpen;
        bool m_writeAccess;

        UInt64 m_creationTime; // Times in ISOM format which is seconds since 1904-01-01
        UInt64 m_modificationTime;
        bool m_valueChanged;

        /// <summary>
        /// Tries to open a file as an ISO Base Media file.
        /// </summary>
        /// <param name="filename">Name of file to attempt to open.</param>
        /// <param name="writeAccess">True if file is to be opened for reading. Defaults to false.</param>
        /// <returns>An IsomCoreMetadata object if successful. Otherwise null.</returns>
        public static IsomCoreMetadata TryOpen(string filename, bool writeAccess = false)
        {
            Stream stream = null;
            try
            {
                stream = new FileStream(filename, FileMode.Open,
                    writeAccess ? FileAccess.ReadWrite : FileAccess.Read,
                    writeAccess ? FileShare.None : FileShare.Read);
                var isom = TryOpen(stream, writeAccess, false);
                if (isom != null)
                {
                    stream = null;
                    return isom;
                }
            }
            catch
            {
                // Do nothing except clean up and return null
            }
            finally
            {
                if (stream != null)
                {
                    stream.Dispose();
                    stream = null;
                }
            }

            return null;
        }

        public static IsomCoreMetadata TryOpen(Stream stream, bool writeAccess = false, bool leaveOpen = false)
        {
            // Attempt to read the header
            stream.Position = 0;
            var size = ReadUInt32BE(stream);
            var boxType = ReadASCII(stream, 4);

            if (size <= stream.Length && boxType.Equals("ftyp", StringComparison.Ordinal))
            {
                return new IsomCoreMetadata(stream, writeAccess, leaveOpen);
            }

            if (!leaveOpen)
            {
                stream.Dispose();
            }

            return null;
        }

        /// <summary>
        /// Create the ISOM object from a filename.
        /// </summary>
        /// <remarks>
        /// .MP4, .MOV, .M4A, and a number of other file formats qualify as ISOM.
        /// </remarks>
        /// <exception cref="InvalidOperationException">Thrown if not an ISOM file format.</exception>
        /// <param name="filename">Name of an ISOM file.</param>
        /// <param name="writeAccess">True if file is to be opened for reading. Defaults to false.</param>
        public IsomCoreMetadata(string filename, bool writeAccess = false)
        {
            Stream stream = null;
            try
            {
                stream = new FileStream(filename, FileMode.Open,
                    writeAccess ? FileAccess.ReadWrite : FileAccess.Read,
                    writeAccess ? FileShare.None : FileShare.Read);
                Init(stream, writeAccess, false);
                stream = null;
            }
            finally
            {
                if (stream != null)
                {
                    stream.Dispose();
                    stream = null;
                }
            }
        }

        /// <summary>
        /// Create the ISOM object from a stream.
        /// </summary>
        /// <param name="stream">The stream containing an ISOM file.</param>
        public IsomCoreMetadata(Stream stream, bool writeAccess = false, bool leaveOpen = false)
        {
            Init(stream, writeAccess, leaveOpen);
        }

        private void Init(Stream stream, bool writeAccess, bool leaveOpen)
        {
            m_stream = stream;
            m_writeAccess = writeAccess;
            m_leaveOpen = leaveOpen;

            DecodeProperties();
            if (!writeAccess)
            {
                Dispose();
            }
        }

        /// <summary>
        /// The MajorBrand field from the ISOM file header. (Read only)
        /// </summary>
        public string MajorBrand { get; private set; }

        /// <summary>
        /// The MinorVersion field from the ISOM file header. (Read only)
        /// </summary>
        public Int32 MinorVersion { get; private set; }

        /// <summary>
        /// The CompatibleBrands field from the ISOM file header. (Read only)
        /// </summary>
        public string[] CompatibleBrands { get; private set; }

        /// <summary>
        /// Play time of the media file. (Read only)
        /// </summary>
        public TimeSpan Duration { get; private set; }

        /// <summary>
        /// Creation date and time. (Read/Write). Null indicates that the value is not set.
        /// </summary>
        public DateTime? CreationTime
        {
            get { return DateTimeFromIsomTime(m_creationTime); }

            set
            {
                if (!m_writeAccess)
                    throw new InvalidOperationException("IsomCoreMetata opened for reading only.");

                m_creationTime = IsomTimeFromDateTime(value);
                m_valueChanged = true;
            }
        }

        /// <summary>
        /// Modification date and time. (Read/Write). Null indicates that the value is not set.
        /// </summary>
        public DateTime? ModificationTime
        {
            get { return DateTimeFromIsomTime(m_modificationTime); }

            set
            {
                if (!m_writeAccess)
                    throw new InvalidOperationException("IsomCoreMetata opened for reading only.");

                m_modificationTime = IsomTimeFromDateTime(value);
                m_valueChanged = true;
            }
        }

        /// <summary>
        /// Save any metadata changes to the file. If the file is closed without committing then changes are not saved.
        /// </summary>
        public void Commit()
        {
            if (!m_writeAccess)
                throw new InvalidOperationException("IsomCoreMetata opened for reading only.");

            if (!m_valueChanged)
                return; // Nothing to write

            UpdateMVHD(new Box(m_stream));
        }

        /// <summary>
        /// Dispose the ISOM file. If changes are made but commit is not called first then changes are lost.
        /// </summary>
        public void Dispose()
        {
            if (m_stream != null)
            {
                if (!m_leaveOpen)
                {
                    m_stream.Dispose();
                }
                m_stream = null;
            }
        }

        #region Private Methods

        void DecodeProperties()
        {
            var root = new Box(m_stream);

            DecodeFTYP(root);
            DecodeMVHD(root);
        }

        void DecodeFTYP(Box root)
        {
            var b = root.FirstChild;

            if (!b.BoxType.Equals("ftyp", StringComparison.Ordinal))
                throw new InvalidOperationException("ISOM file does not start with an 'ftyp' record.");
            m_stream.Position = (long)b.Body;
            MajorBrand = ReadASCII(m_stream, 4);
            MinorVersion = (int)ReadUInt32BE(m_stream);

            var comps = new List<string>();
            while (m_stream.Position <= (long)(b.BodyEnd - 4))
            {
                comps.Add(ReadASCII(m_stream, 4));
            }
            CompatibleBrands = comps.ToArray();
        }

        void DecodeMVHD(Box root)
        {
            var moov = root.FindChild("moov");
            if (moov == null) throw new InvalidOperationException("ISOM file does not have the required 'moov' record.");

            var mvhd = moov.FindChild("mvhd");
            if (mvhd == null) throw new InvalidOperationException("ISOM file does not have the required 'mvhd' record.");

            m_stream.Position = (long)mvhd.Body;
            var version = ReadByte(m_stream);
            m_stream.Seek(3, SeekOrigin.Current); // Flags are unused

            ulong timescale;
            ulong duration;
            if (version == 1)
            {
                m_creationTime = ReadUInt64BE(m_stream);
                m_modificationTime = ReadUInt64BE(m_stream);
                timescale = ReadUInt64BE(m_stream);
                duration = ReadUInt64BE(m_stream);
            }
            else // version == 0
            {
                m_creationTime = ReadUInt32BE(m_stream);
                m_modificationTime = ReadUInt32BE(m_stream);
                timescale = ReadUInt32BE(m_stream);
                duration = ReadUInt32BE(m_stream);
            }

            Duration = TimeSpanFromIsom(duration, timescale);
        }

        void UpdateMVHD(Box root)
        {
            var moov = root.FindChild("moov");
            if (moov == null) throw new InvalidOperationException("ISOM file does not have the required 'moov' record.");

            var mvhd = moov.FindChild("mvhd");
            if (mvhd == null) throw new InvalidOperationException("ISOM file does not have the required 'mvhd' record.");

            m_stream.Position = (long)mvhd.Body;
            var version = ReadByte(m_stream);
            m_stream.Seek(3, SeekOrigin.Current); // Flags are unused

            if (version == 1)
            {
                WriteUInt64BE(m_stream, m_creationTime);
                WriteUInt64BE(m_stream, m_modificationTime);
            }
            else // version == 0 (this will overflow on 2040-02-06)
            {
                WriteUInt32BE(m_stream, (UInt32)m_creationTime);
                WriteUInt32BE(m_stream, (UInt32)m_modificationTime);
            }
            m_stream.Flush();
        }

        #endregion // Private Methods

        /// <summary>
        /// ISOM files are composed of boxes. Represents one box.
        /// </summary>
        private class Box
        {
            private readonly Stream m_stream;

            /// <summary>
            ///     Creates the root box
            /// </summary>
            /// <param name="stream">The Stream for the file.</param>
            public Box(Stream stream)
            {
                m_stream = stream;
                Parent = null;
                Start = 0;
                Size = (ulong)stream.Length;
                Body = 0;
                BoxType = string.Empty;
            }

            private Box(Box parent, ulong start)
            {
                m_stream = parent.m_stream;
                Parent = parent;
                Start = start;
                m_stream.Position = (long)start;
                var size = ReadUInt32BE(m_stream);
                BoxType = ReadASCII(m_stream, 4);
                Body = Start + 8;
                if (size == 0)
                {
                    Size = (ulong)m_stream.Length - Start;
                }
                else if (size == 1)
                {
                    Size = ReadUInt64BE(m_stream);
                    Body += 8;
                }
                else
                {
                    Size = size;
                }
                if (BoxType.Equals("uuid", StringComparison.Ordinal))
                {
                    var uuid = new Guid(ReadBytes(m_stream, 16));
                    BoxType = uuid.ToString();
                    Body += 16;
                }
            }

            private ulong Start { get; }

            private ulong Size { get; }

            public string BoxType { get; }

            public ulong Body { get; }

            public ulong BodyEnd => (Start + Size);

            private Box Parent { get; }

            public Box FirstChild => new Box(this, Body);

            public Box NextSibling
            {
                get
                {
                    var start = Start + Size;
                    if (start >= Parent.Start + Parent.Size)
                    {
                        return null;
                    }
                    return new Box(Parent, Start + Size);
                }
            }

            public Box FindChild(string type)
            {
                for (var b = FirstChild; b != null; b = b.NextSibling)
                {
                    if (b.BoxType.Equals(type, StringComparison.Ordinal))
                    {
                        return b;
                    }
                }
                return null;
            }
        }

        #region DateTime to ISOM conversions

        static readonly DateTime s_isomBaseTime = new DateTime(1904, 1, 1, 0, 0, 0, DateTimeKind.Utc);

        static UInt64 IsomTimeFromDateTime(DateTime? value)
        {
            if (value == null)
            {
                return 0;
            }

            DateTime v = (DateTime)value;
            // Assume UTC if kind is not specified
            if (v.Kind == DateTimeKind.Unspecified)
            {
                v = DateTime.SpecifyKind(v, DateTimeKind.Utc);
            }
            else if (v.Kind == DateTimeKind.Local)
            {
                v = v.ToUniversalTime();
            }

            return (UInt64)v.Subtract(s_isomBaseTime).TotalSeconds;
        }

        static DateTime? DateTimeFromIsomTime(UInt64 value)
        {
            if (value == 0)
            {
                return null;
            }

            return s_isomBaseTime.AddSeconds(value);
        }

        static TimeSpan TimeSpanFromIsom(UInt64 duration, UInt64 timescale)
        {
            if (duration == ulong.MaxValue)
            {
                return TimeSpan.FromTicks(0);
            }

            UInt64 multiplier = 10000000L;

            // Avoid overflow with large timescale values
            while (multiplier >= 10 && (timescale % 10) == 0)
            {
                multiplier /= 10;
                timescale /= 10;
            }

            return TimeSpan.FromTicks((long)(duration * multiplier / timescale));
        }

        #endregion

        #region Binary helper functions

        static byte ReadByte(Stream stream)
        {
            int b = stream.ReadByte();
            if (b < 0)
            {
                throw new InvalidOperationException("Unexpected end of ISOM file.");
            }
            return (byte)b;
        }

        static byte[] ReadBytes(Stream stream, int length)
        {
            byte[] buf = new byte[length];
            var bytesRead = stream.Read(buf, 0, length);
            if (bytesRead < length)
            {
                throw new InvalidOperationException("Unexpected end of ISOM file.");
            }
            return buf;
        }

        static UInt32 ReadUInt32BE(Stream stream)
        {
            var a = ReadBytes(stream, 4);
            Array.Reverse(a);
            return BitConverter.ToUInt32(a, 0);
        }

        static UInt64 ReadUInt64BE(Stream stream)
        {
            var a = ReadBytes(stream, 8);
            Array.Reverse(a);
            return BitConverter.ToUInt64(a, 0);
        }

        static string ReadASCII(Stream stream, int length)
        {
            return Encoding.ASCII.GetString(ReadBytes(stream, length));
        }

        static void WriteUInt32BE(Stream stream, UInt32 value)
        {
            var a = BitConverter.GetBytes(value);
            Array.Reverse(a);
            stream.Write(a, 0, a.Length);
        }

        static void WriteUInt64BE(Stream stream, UInt64 value)
        {
            var a = BitConverter.GetBytes(value);
            Array.Reverse(a);
            stream.Write(a, 0, a.Length);
        }

        #endregion // Binary read helper functions

    }
}