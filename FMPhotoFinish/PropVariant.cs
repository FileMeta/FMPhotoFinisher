/*
---
name: PropVariant.cs
description: Managed Code Interoperability for PROPVARIANT - extends existing support for VARIANT
url: https://github.com/FileMeta/WinShellPropertyStore/raw/master/PropVariant.cs
version: 1.3
keywords: CodeBit
dateModified: 2019-05-10
license: http://unlicense.org
# Metadata in MicroYaml format. See http://filemeta.org and http://schema.org
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
   Help in managing VARIANT from managed code: https://limbioliong.wordpress.com/2011/09/04/using-variants-in-managed-code-part-1/
*/

using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Interop
{

    internal static class PropVariant
    {
        /* C++ version
        typedef struct PROPVARIANT {
            VARTYPE vt;
            WORD    wReserved1;
            WORD    wReserved2;
            WORD    wReserved3;
            union {
                // Various types of up to 8 bytes in 32-bit mode and 16 bytes in 64-bit mode
            } data;
        } PROPVARIANT;
        */

        public static readonly int sizeofPROPVARIANT = 8 + 2 * IntPtr.Size;
        const int offsetofPROPVARIANT_vt = 0;
        const int offsetofPROPVARIANT_data = 8;
        const int offsetofPROPVARIANT_cElems = 8;
        static readonly int offsetofPROPVARIANT_pElems = 8 + IntPtr.Size;

        /// <summary>
        /// Initialize a VARIANT or PROPVARIANT structure to all zeros.
        /// </summary>
        /// <param name="pVariant">Address of the variant to initialize.</param>
        public static void PropVariantInit(IntPtr pVariant)
        {
            Marshal.WriteInt64(pVariant, 0);
            Marshal.WriteInt64(pVariant, 8, 0);
            if (sizeofPROPVARIANT > 16)
            {
                Marshal.WriteInt64(pVariant, 16, 0);
            }
        }

        /// <summary>
        /// Initialize a VARIANT or PROPVARIANT structure to a vt with zero data
        /// </summary>
        /// <param name="pVariant">Address of the variant to initialize.</param>
        private static void PropVariantInit(IntPtr pVariant, VT vt)
        {
            Marshal.WriteInt64(pVariant, (Int16)vt);    // Writes the VT into the first 16 bits and clears the remaining 48 bits to zero
            Marshal.WriteInt64(pVariant, 8, 0);
            if (sizeofPROPVARIANT > 16)
            {
                Marshal.WriteInt64(pVariant, 16, 0);
            }
        }


        /// <summary>
        /// Converts a string to a PropVariant with type LPSTR instead of BSTR.
        /// </summary>
        /// <param name="value">The string to convert.</param>
        /// <remarks>
        /// <para>Unlike <see cref="Marshal.GetNativeVariantForObject"/>, this method will
        /// use VT_LPWSTR for storing the string rather that VT_BSTR. The storage for the string
        /// is allocated using IMalloc-compatible methods.</para>
        /// </remarks>
        public static void GetPropVariantLPWSTR(string value, IntPtr pVariant)
        {
            PropVariantInit(pVariant, VT.LPWSTR);
            Marshal.WriteIntPtr(pVariant, offsetofPROPVARIANT_data, Marshal.StringToCoTaskMemUni(value));
        }

        /// <summary>
        /// Converts a <see cref="DateTime"/> to a PropVariant with type FILETIME instead of DATETIME.
        /// </summary>
        /// <param name="value">The DateTime to convert.</param>
        /// <remarks>
        /// <para>Unlike <see cref="Marshal.GetNativeVariantForObject"/>, this method will
        /// use VT_FILETIME for storing the date rather that VT_DATE.</para>
        /// </remarks>
        public static void GetPropVariantFILETIME(DateTime value, IntPtr pVariant)
        {
            PropVariantInit(pVariant, VT.FILETIME);
            Marshal.WriteInt64(pVariant, offsetofPROPVARIANT_data, value.ToFileTimeUtc());
        }

        /// <summary>
        /// Converts a managed object to unmanaged variant using LPWSTR for strings and FILETIME for DateTime
        /// </summary>
        /// <param name="value">The value to convert</param>
        /// <param name="pv">Location of the variant to load with the object</param>
        public static void GetPropVariantFromObject(object value, IntPtr pv)
        {
            switch (value)
            {
                case string str:
                    GetPropVariantLPWSTR(str, pv);
                    break;

                case DateTime dt:
                    GetPropVariantFILETIME(dt, pv);
                    break;

                default:
                    Marshal.GetNativeVariantForObject(value, pv);
                    break;
            }
        }

        public static object GetObjectFromPropVariant(IntPtr pv)
        {
            var vt = (VT)Marshal.ReadInt16(pv);

            object value = null;
            switch (vt)
            {
                case VT.EMPTY:
                case VT.NULL:
                case VT.I2:
                case VT.I4:
                case VT.R4:
                case VT.R8:
                case VT.CY:
                case VT.DATE:
                case VT.BSTR:
                case VT.ERROR:
                case VT.BOOL:
                case VT.DECIMAL:
                case VT.I1:
                case VT.UI1:
                case VT.UI2:
                case VT.UI4:
                case VT.I8:
                case VT.UI8:
                case VT.INT:
                case VT.UINT:
                case VT.VOID:
                case VT.HRESULT:
                    value = Marshal.GetObjectForNativeVariant(pv);
                    break;

                case VT.LPSTR:
                    value = Marshal.PtrToStringAnsi(Marshal.ReadIntPtr(pv, offsetofPROPVARIANT_data));
                    break;

                case VT.LPWSTR:
                    value = Marshal.PtrToStringUni(Marshal.ReadIntPtr(pv, offsetofPROPVARIANT_data));
                    break;

                case VT.FILETIME:
                    value = DateTime.FromFileTimeUtc(Marshal.ReadInt64(pv, offsetofPROPVARIANT_data));
                    break;

                case VT.CLSID:
                    value = Marshal.PtrToStructure<Guid>(Marshal.ReadIntPtr(pv, offsetofPROPVARIANT_data));
                    break;

                case VT.VECTOR|VT.I2:
                    {
                        int cElems = (int)Marshal.ReadIntPtr(pv, offsetofPROPVARIANT_cElems);
                        Int16[] a = new Int16[cElems];
                        Marshal.Copy(Marshal.ReadIntPtr(pv, offsetofPROPVARIANT_pElems), a, 0, cElems);
                        value = a;
                    }
                    break;

                case VT.VECTOR|VT.I4:
                case VT.VECTOR|VT.INT:
                    {
                        int cElems = (int)Marshal.ReadIntPtr(pv, offsetofPROPVARIANT_cElems);
                        Int32[] a = new Int32[cElems];
                        Marshal.Copy(Marshal.ReadIntPtr(pv, offsetofPROPVARIANT_pElems), a, 0, cElems);
                        value = a;
                    }
                    break;

                case VT.VECTOR|VT.R4:
                    {
                        int cElems = (int)Marshal.ReadIntPtr(pv, offsetofPROPVARIANT_cElems);
                        float[] a = new float[cElems];
                        Marshal.Copy(Marshal.ReadIntPtr(pv, offsetofPROPVARIANT_pElems), a, 0, cElems);
                        value = a;
                    }
                    break;

                case VT.VECTOR|VT.I1:
                    {
                        int cElems = (int)Marshal.ReadIntPtr(pv, offsetofPROPVARIANT_cElems);
                        byte[] a = new byte[cElems];
                        Marshal.Copy(Marshal.ReadIntPtr(pv, offsetofPROPVARIANT_pElems), a, 0, cElems);
                        SByte[] b = new SByte[cElems];
                        Buffer.BlockCopy(a, 0, b, 0, cElems);
                        value = b;
                    }
                    break;

                case VT.VECTOR|VT.UI1:
                    {
                        int cElems = (int)Marshal.ReadIntPtr(pv, offsetofPROPVARIANT_cElems);
                        byte[] a = new byte[cElems];
                        Marshal.Copy(Marshal.ReadIntPtr(pv, offsetofPROPVARIANT_pElems), a, 0, cElems);
                    }
                    break;

                case VT.VECTOR|VT.UI2:
                    {
                        int cElems = (int)Marshal.ReadIntPtr(pv, offsetofPROPVARIANT_cElems);
                        Int16[] a = new Int16[cElems];
                        Marshal.Copy(Marshal.ReadIntPtr(pv, offsetofPROPVARIANT_pElems), a, 0, cElems);
                        UInt16[] b = new UInt16[cElems];
                        Buffer.BlockCopy(a, 0, b, 0, cElems * 2);
                        value = b;
                    }
                    break;

                case VT.VECTOR|VT.UI4:
                case VT.VECTOR|VT.UINT:
                    {
                        int cElems = (int)Marshal.ReadIntPtr(pv, offsetofPROPVARIANT_cElems);
                        Int32[] a = new Int32[cElems];
                        Marshal.Copy(Marshal.ReadIntPtr(pv, offsetofPROPVARIANT_pElems), a, 0, cElems);
                        UInt32[] b = new UInt32[cElems];
                        Buffer.BlockCopy(a, 0, b, 0, (int)cElems * 4);
                        value = b;
                    }
                    break;

                case VT.VECTOR|VT.I8:
                    {
                        int cElems = (int)Marshal.ReadIntPtr(pv, offsetofPROPVARIANT_cElems);
                        Int64[] a = new Int64[cElems];
                        Marshal.Copy(Marshal.ReadIntPtr(pv, offsetofPROPVARIANT_pElems), a, 0, cElems);
                        value = a;
                    }
                    break;

                case VT.VECTOR|VT.UI8:
                    {
                        int cElems = (int)Marshal.ReadIntPtr(pv, offsetofPROPVARIANT_cElems);
                        Int64[] a = new Int64[cElems];
                        Marshal.Copy(Marshal.ReadIntPtr(pv, offsetofPROPVARIANT_pElems), a, 0, cElems);
                        UInt64[] b = new UInt64[cElems];
                        Buffer.BlockCopy(a, 0, b, 0, cElems * 8);
                        value = b;
                    }
                    break;

                case VT.VECTOR|VT.R8:
                    {
                        int cElems = (int)Marshal.ReadIntPtr(pv, offsetofPROPVARIANT_cElems);
                        double[] a = new double[cElems];
                        Marshal.Copy(Marshal.ReadIntPtr(pv, offsetofPROPVARIANT_pElems), a, 0, cElems);
                        value = a;
                    }
                    break;

                case VT.VECTOR|VT.LPSTR:
                    {
                        int cElems = (int)Marshal.ReadIntPtr(pv, offsetofPROPVARIANT_cElems);
                        IntPtr pElems = Marshal.ReadIntPtr(pv, offsetofPROPVARIANT_pElems);
                        string[] strings = new string[cElems];
                        for (int i = 0; i < cElems; ++i)
                        {
                            strings[i] = Marshal.PtrToStringAnsi(Marshal.ReadIntPtr(pElems, i * IntPtr.Size));
                        }
                        value = strings;
                    }
                    break;

                case VT.VECTOR|VT.LPWSTR:
                    {
                        int cElems = (int)Marshal.ReadIntPtr(pv, offsetofPROPVARIANT_cElems);
                        IntPtr pElems = Marshal.ReadIntPtr(pv, offsetofPROPVARIANT_pElems);
                        string[] strings = new string[cElems];
                        for (int i = 0; i < cElems; ++i)
                        {
                            strings[i] = Marshal.PtrToStringUni(Marshal.ReadIntPtr(pElems, i * IntPtr.Size));
                        }
                        value = strings;
                    }
                    break;

                default:
#if DEBUG
                    // Attempt conversion and report if it works
                    try
                    {
                        value = Marshal.GetObjectForNativeVariant(pv);
                        if (value == null) value = "(null)";
                        value = String.Format("(Supported type 0x{0:x4}): {1}", vt, value.ToString());
                    }
                    catch
                    {
                        throw new NotImplementedException(string.Format($"Conversion of PROPVARIANT type VT_{vt} is not yet supported."));
                    }

                    throw new NotImplementedException(string.Format($"Conversion of PROPVARIANT type VT_{vt} is not yet supported but Marshal.GetObjectForNativeVariant seems to work."));
#else
                    throw new NotImplementedException(string.Format($"Conversion of PROPVARIANT type VT_{vt} is not yet supported."));
#endif
            }

            return value;
        } // Method PropVariantToObject

        /// <summary>
        /// Returns the managed type that will result from conversion of a variant type
        /// </summary>
        /// <param name="vt">The variant type to be converted. (May be coerced from
        /// int or uint.)</param>
        /// <returns>The managed type that would result from conversion of the variant
        /// type.</returns>
        /// <remarks>
        /// <para>VT_EMPTY and VT_NULL return typeof(object). Conversion of either of
        /// these types would result in null.
        /// </para>
        /// <para>Unsupported VT values result in a null return.
        /// </para>
        /// </remarks>
        public static Type GetManagedTypeFromVariantType(VT vt)
        {
            switch (vt)
            {
                case VT.EMPTY: // 0
                case VT.NULL:  // 1
                    return typeof(object);
                case VT.I2: // 2
                    return typeof(Int16);
                case VT.I4: // 3
                case VT.INT: // 22
                    return typeof(Int32);
                case VT.R4: // 4
                    return typeof(float);
                case VT.R8: // 5
                    return typeof(double);
                case VT.CY: // 6
                    return typeof(decimal);
                case VT.DATE: // 7
                    return typeof(DateTime);
                case VT.BSTR: // 8
                case VT.LPSTR: // 30
                case VT.LPWSTR: // 31
                    return typeof(string);
                case VT.ERROR: // 10
                    return typeof(UInt32);
                case VT.BOOL: // 11
                    return typeof(bool);
                case VT.DECIMAL: // 14
                    return typeof(decimal);
                case VT.I1: // 16
                    return typeof(sbyte);
                case VT.UI1: // 17
                    return typeof(byte);
                case VT.UI2: // 18
                    return typeof(UInt16);
                case VT.UI4: // 19
                case VT.UINT: // 23
                case VT.HRESULT: // 25
                    return typeof(UInt32);
                case VT.I8: // 20
                    return typeof(Int64);
                case VT.UI8: // 21
                    return typeof(UInt64);
                case VT.FILETIME: // 64
                    return typeof(DateTime);
                case VT.CLSID: // 72
                    return typeof(Guid);
                case VT.VECTOR | VT.I2:
                    return typeof(Int16[]);
                case VT.VECTOR | VT.I4:
                case VT.VECTOR | VT.INT:
                    return typeof(Int32[]);
                case VT.VECTOR | VT.R4:
                    return typeof(float[]);
                case VT.VECTOR | VT.R8:
                    return typeof(double[]);
                case VT.VECTOR | VT.I1:
                    return typeof(sbyte[]);
                case VT.VECTOR | VT.UI1:
                    return typeof(byte[]);
                case VT.VECTOR | VT.UI2:
                    return typeof(UInt16[]);
                case VT.VECTOR | VT.UI4:
                case VT.VECTOR | VT.UINT:
                    return typeof(UInt32[]);
                case VT.VECTOR | VT.I8:
                    return typeof(Int64[]);
                case VT.VECTOR | VT.UI8:
                    return typeof(UInt64[]);
                case VT.VECTOR | VT.LPSTR:
                case VT.VECTOR | VT.LPWSTR:
                    return typeof(String[]);

                // Explicitly unsupported types
                case VT.DISPATCH:
                case VT.UNKNOWN:
                case VT.STREAM: // 0x42 VT_STREAM (Used by: System.ThumbnailStream on .mp3 format file, System.Contact.AccountPictureLarge, System.Contact.AccountPictureSmall)
                case VT.CF: // 0x47 VT_CF (CLIPDATA format, Used by: System.Thumbnail on .xls format file)
                case VT.VECTOR | VT.VARIANT: // VT_VECTOR|VT_VARIANT (Used by: Unnamed property on .potx file)
                    return null;

                default:
#if DEBUG
                        throw new NotImplementedException(string.Format("Unexpected PROPVARIANT type 0x{0:x4}.", (uint)vt));
#else
                    return null;
#endif
            }
        }

        // Reference: https://msdn.microsoft.com/en-us/library/windows/desktop/aa380072(v=vs.85).aspx

        [DllImport(@"ole32.dll", SetLastError = true, CallingConvention = CallingConvention.StdCall, PreserveSig = false)]
        public static extern void PropVariantClear([In] IntPtr pvarg);

        public enum VT : Int16
        {
            EMPTY = 0,
            NULL = 1,
            I2 = 2,
            I4 = 3,
            R4 = 4,
            R8 = 5,
            CY = 6,
            DATE = 7,   // 64-bit floating point, number of days since December 31, 1899
            BSTR = 8,   // Traditional way to pass strings
            DISPATCH = 9,   // IDispatch interface to object
            ERROR = 10, // 4 bytes (DWORD) containing status code
            BOOL = 11,
            VARIANT = 12,   // For combining with VECTOR or BYREF
            UNKNOWN = 13,   // IUnknown interface to object
            DECIMAL = 14,
            I1 = 16,    // Signed byte
            UI1 = 17,    // Unsigned byte
            UI2 = 18,
            UI4 = 19,
            I8 = 20,
            UI8 = 21,

            INT = 22,   // 4 bytes, equivalent to I4
            UINT = 23,  // 4 bytes, equivalent to UI4
            VOID = 24,
            HRESULT = 25,

            LPSTR = 30,
            LPWSTR = 31,

            FILETIME = 64,
            BLOB = 65,
            STREAM = 66,
            STORAGE = 67,   // IStorage interface
            STREAMED_OBJECT = 68,
            STORED_OBJECT = 69,  // IStorage containing a "loadable object"
            BLOBOBJECT = 70,
            CF = 71,    // Pointer to CLIPDATA structure
            CLSID = 72,
            VERSIONED_STREAM = 73,

            // The following are modifiers that may be OR'ed to certain types
            VECTOR     = 0x1000,
            ARRAY      = 0x2000,
            BYREF      = 0x4000,

            // The following can be used to mask off the type when VECTOR, ARRAY, OR BYREF is used
            TYPEMASK   = 0x0FFF
        }

    } // class PropVariant

} // namespace WinShell
