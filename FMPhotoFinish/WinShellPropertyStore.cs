/*
---
name: WinShellPropertyStore.cs
description: C# Wrapper for Windows Property System
url: https://github.com/FileMeta/WinShellPropertyStore/raw/master/WinShellPropertyStore.cs
version: 1.6
keywords: CodeBit
dateModified: 2018-12-13
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
   Windows Property System: https://msdn.microsoft.com/en-us/library/windows/desktop/ff728898(v=vs.85).aspx
   Help in managing VARIANT from managed code: https://limbioliong.wordpress.com/2011/09/04/using-variants-in-managed-code-part-1/
*/

// This interface layer corrects certain problems with the Windows Property Store. To suppress
// the corrections, uncomment the following line.
//#define RAW_PROPERTY_STORE

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace WinShell
{

    /// <summary>
    /// Wrapper class for Windows Shell IPropertyStore. Supports read and
    /// write access to metadata properties on files using the Windows
    /// Property System.
    /// </summary>
    /// <remarks>
    /// Inherits from IDisposable. You should always use "using" statements
    /// with this class. If properties are written, be sure to call "commit"
    /// before the instance is disposed.
    /// </remarks>
    class PropertyStore : IDisposable
    {

#if !RAW_PROPERTY_STORE
        static PROPERTYKEY s_pkContentType = new PROPERTYKEY("D5CDD502-2E9C-101B-9397-08002B2CF9AE", 26); // System.ContentType
        static PROPERTYKEY s_pkItemDate = new PROPERTYKEY("f7db74b4-4287-4103-afba-f1b13dcd75cf", 100); // System.ItemDate
        static PROPERTYKEY s_pkDateTaken = new PROPERTYKEY("14B81DA1-0135-4D31-96D9-6CBFC9671A99", 36867);
        static PROPERTYKEY s_pkDateEncoded = new PROPERTYKEY("2e4b640d-5019-46d8-8881-55414cc5caa0", 100); // System.Media.DateEncoded
#endif

        /// <summary>
        /// Open the property store for a file.
        /// </summary>
        /// <param name="filename">Path or filename of file on which to open the property store.</param>
        /// <param name="writeAccess">True if write access is desired. Defaults to false.</param>
        /// <returns>A <see cref="PropertyStore"/> instance.</returns>
        public static PropertyStore Open(string filename, bool writeAccess = false)
        {
            PropertyStoreInterop.IPropertyStore store;
            Guid iPropertyStoreGuid = typeof(PropertyStoreInterop.IPropertyStore).GUID;
            PropertyStoreInterop.SHGetPropertyStoreFromParsingName(filename, IntPtr.Zero,
                writeAccess ? PropertyStoreInterop.GETPROPERTYSTOREFLAGS.GPS_READWRITE : PropertyStoreInterop.GETPROPERTYSTOREFLAGS.GPS_BESTEFFORT,
                ref iPropertyStoreGuid, out store);
            return new PropertyStore(store);
        }

        private PropertyStoreInterop.IPropertyStore m_IPropertyStore;
        private PropertyStoreInterop.IPropertyStoreCapabilities m_IPropertyStoreCapabilities;
#if !RAW_PROPERTY_STORE
        private string m_contentType;
#endif

        private PropertyStore(PropertyStoreInterop.IPropertyStore propertyStore)
        {
            m_IPropertyStore = propertyStore;
#if !RAW_PROPERTY_STORE
            m_contentType = GetValue(s_pkContentType) as string;
#endif
        }

        /// <summary>
        /// Gets the number of properties attached to the file.
        /// </summary>
        public int Count
        {
            get
            {
                uint value;
                m_IPropertyStore.GetCount(out value);
                return (int)value;
            }
        }

        /// <summary>
        /// Gets a property key from an item's array of properties.
        /// </summary>
        /// <param name="index">The index of the property key in the property store's
        /// array of <see cref="PROPERTYKEY"/> structures. This is a zero-based index.</param>
        /// <returns>The <see cref="PROPERTYKEY"/> at the specified index.</returns>
        /// <remarks>
        /// This call, in combination with the <see cref="Count"/> property can be
        /// used to enumerate all properties in the PropertyStore.
        /// </remarks>
        public PROPERTYKEY GetAt(int index)
        {
            PROPERTYKEY key;
            m_IPropertyStore.GetAt((uint)index, out key);
            return key;
        }

        /// <summary>
        /// Gets data for a specific property.
        /// </summary>
        /// <param name="key">The <see cref="PROPERTYKEY"/> of the property to be retrieved.</param>
        /// <returns>The data for the specified property or null if the item does not have the property.</returns>
        public object GetValue(PROPERTYKEY key)
        {
            IntPtr pv = IntPtr.Zero;
            object value = null;
            try
            {
                pv = Marshal.AllocCoTaskMem(32); // Structure is 16 bytes in 32-bit windows and 24 bytes in 64-bit but we leave a few more bytes for good measure.
                m_IPropertyStore.GetValue(key, pv);
                value = PropertyStoreInterop.PropVariantToObject(pv);

#if !RAW_PROPERTY_STORE
                // PropertyStore returns all DateTimes in UTC. However, DateTaken certain fields are stored in local time.
                // The conversion to UTC will not be correct unless the timezone on the camera, when the
                // photo was taken, is the same as the timezone on the computer when the value is retrieved.
                // We fix this up by converting back to local time using the computer's current timezone
                // (the same timezone that was used to convert to UTC moments earlier).
                // This works well because unlike the source FILETIME (which should always be UTC), The
                // managed DateTime format has a property that indicates whether the time is local or UTC.
                // Callers should still take care to interpret the time as local to where the photo was taken
                // and not local to where the computer is at present.
                if (value != null
                    && ((string.Equals(m_contentType, "image/jpeg")
                        && (key.Equals(s_pkItemDate) || key.Equals(s_pkDateTaken)))
                    || (string.Equals(m_contentType, "video/avi")
                        && (key.Equals(s_pkItemDate) || key.Equals(s_pkDateEncoded)))))
                {
                    DateTime dt = (DateTime)value;
                    Debug.Assert(dt.Kind == DateTimeKind.Utc);
                    value = dt.ToLocalTime();
                }
#endif
            }
            finally
            {
                if (pv != IntPtr.Zero)
                {
                    try
                    {
                        PropertyStoreInterop.PropVariantClear(pv);
                    }
                    catch
                    {
                        Debug.Fail("VariantClear failure");
                    }
                    Marshal.FreeCoTaskMem(pv);
                    pv = IntPtr.Zero;
                }
            }
            return value;
        }

        /// <summary>
        /// Sets a new property value, or replaces or removes an existing value.
        /// </summary>
        /// <param name="key">The <see cref="PROPERTYKEY"/> of the property to be set.</param>
        /// <param name="value">The value to be set. Managed code values are converted
        /// to the appropriate PROPVARIANT values.</param>
        /// <remarks>
        /// <para>The property store must be opened with write access. (See <see cref="Open"/>).
        /// </para>
        /// <para>Subsequent calls to <see cref="GetCount"/> and <see cref="GetValue"/> will
        /// reflect the new property.
        /// </para>
        /// <para>No changes are made to the underlying file until <see cref="Commit"/> is
        /// called. If Commit is not called then the changes are discarded.
        /// </para>
        /// <para>Removal of a property is not supported by the Windows Property System.
        /// </para>
        /// </remarks>
        public void SetValue(PROPERTYKEY key, object value)
        {
            IntPtr pv = IntPtr.Zero;
            try
            {
                if (value is DateTime)
                {
                    DateTime dt = (DateTime)value;
                    value = dt.ToUniversalTime();
                }

                pv = PropertyStoreInterop.PropVariantFromObject(value);
                m_IPropertyStore.SetValue(key, pv);
            }
            finally
            {
                if (pv != IntPtr.Zero)
                {
                    PropertyStoreInterop.PropVariantClear(pv);
                    Marshal.FreeCoTaskMem(pv);
                    pv = IntPtr.Zero;
                }
            }
        }

        public bool IsPropertyWriteable(PROPERTYKEY key)
        {
            if (m_IPropertyStoreCapabilities == null)
            {
                m_IPropertyStoreCapabilities = (PropertyStoreInterop.IPropertyStoreCapabilities)m_IPropertyStore;
            }

            Int32 hResult = m_IPropertyStoreCapabilities.IsPropertyWritable(ref key);
            if (hResult == 0) return true;
            if (hResult > 0) return false;
            Marshal.ThrowExceptionForHR(hResult);
            return false;   // This should not occur
        }

        /// <summary>
        /// Saves a set of property changes to the underlying file.
        /// </summary>
        public void Commit()
        {
            m_IPropertyStore.Commit();
        }

        /// <summary>
        /// Closes the property store and releases all associated resources.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
        }

        ~PropertyStore()
        {
            Dispose(false);
        }

        private void Dispose(bool disposing)
        {
            if (m_IPropertyStore != null)
            {
                if (!disposing)
                {
                    Debug.Fail("Failed to dispose PropertyStore");
                }

                Marshal.FinalReleaseComObject(m_IPropertyStore);
                m_IPropertyStore = null;
            }

            if (m_IPropertyStoreCapabilities != null)
            {
                Marshal.FinalReleaseComObject(m_IPropertyStoreCapabilities);
                m_IPropertyStoreCapabilities = null;
            }

            if (disposing)
            {
                GC.SuppressFinalize(this);
            }
        }

    } // class PropertyStore

    /// <summary>
    /// Wrapper class for Windows Shell IPropertySystem. Provides classes for
    /// looking up property IDs and translating between property IDs and
    /// property names.
    /// </summary>
    public class PropertySystem : IDisposable
    {
        PropertyStoreInterop.IPropertySystem m_IPropertySystem;

        /// <summary>
        /// Constructs an instance of the PropertySytem object. PropertySystem has no
        /// state so one instance can be shared across the whole application if that is
        /// convenient.
        /// </summary>
        /// <remarks>
        /// PropertySystem should be disposed before going out of scope in order to free
        /// associated system resources.
        /// </remarks>
        public PropertySystem()
        {
            Guid IID_IPropertySystem = typeof(PropertyStoreInterop.IPropertySystem).GUID;
            PropertyStoreInterop.PSGetPropertySystem(ref IID_IPropertySystem, out m_IPropertySystem);
        }

        /// <summary>
        /// Get the <see cref="PropertyDescription"/> for a particular <see cref="PROPERTYKEY"/>.
        /// </summary>
        /// <param name="propKey">The <see cref="PROPERTYKEY"/> for which the description is to be retrieved.</param>
        /// <returns>A <see cref="PropertyDescription"/>. Null if the PROPERTYKEY does not have a registered description.</returns>
        public PropertyDescription GetPropertyDescription(PROPERTYKEY propKey)
        {
            Int32 hResult;
            Guid IID_IPropertyDescription = typeof(PropertyStoreInterop.IPropertyDescription).GUID;
            PropertyStoreInterop.IPropertyDescription iPropertyDescription = null;
            try
            {
                hResult = m_IPropertySystem.GetPropertyDescription(propKey, ref IID_IPropertyDescription, out iPropertyDescription);
                if (hResult < 0) return null;
                return new PropertyDescription(iPropertyDescription);
            }
            finally
            {
                if (iPropertyDescription != null)
                {
                    Marshal.FinalReleaseComObject(iPropertyDescription);
                    iPropertyDescription = null;
                }
            }
        }

        /// <summary>
        /// Get the <see cref="PropertyDescription"/> that corresponds to a name.
        /// </summary>
        /// <param name="canonicalName">The canonical name of the property to be retrieved.</param>
        /// <returns>The <see cref="PropertyDescription"/> that corresponds to the specified name.</returns>
        /// <remarks>
        /// <para>Returns null if no property matches the name.
        /// </para>
        /// <para>Canonical names for properties defined by Windows are listed at
        /// <see cref="https://msdn.microsoft.com/en-us/library/windows/desktop/dd561977(v=vs.85).aspx"/></para>
        /// </remarks>
        public PropertyDescription GetPropertyDescriptionByName(string canonicalName)
        {
            Int32 hResult;
            Guid IID_IPropertyDescription = typeof(PropertyStoreInterop.IPropertyDescription).GUID;
            PropertyStoreInterop.IPropertyDescription iPropertyDescription = null;
            try
            {
                hResult = m_IPropertySystem.GetPropertyDescriptionByName(canonicalName, ref IID_IPropertyDescription, out iPropertyDescription);
                if (hResult < 0)
                {
                    return null;
                }
                return new PropertyDescription(iPropertyDescription);
            }
            finally
            {
                if (iPropertyDescription != null)
                {
                    Marshal.FinalReleaseComObject(iPropertyDescription);
                    iPropertyDescription = null;
                }
            }
        }

        public PROPERTYKEY GetPropertyKeyByName(string canonicalName)
        {
            Int32 hResult;
            PROPERTYKEY propertyKey;    // Initializes automatically to all zeros

            Guid IID_IPropertyDescription = typeof(PropertyStoreInterop.IPropertyDescription).GUID;
            PropertyStoreInterop.IPropertyDescription iPropertyDescription = null;
            try
            {
                hResult = m_IPropertySystem.GetPropertyDescriptionByName(canonicalName, ref IID_IPropertyDescription, out iPropertyDescription);
                if (hResult < 0)
                {
                    Marshal.ThrowExceptionForHR(hResult);
                }
                hResult = iPropertyDescription.GetPropertyKey(out propertyKey);
                if (hResult < 0)
                {
                    Marshal.ThrowExceptionForHR(hResult);
                }
            }
            finally
            {
                if (iPropertyDescription != null)
                {
                    Marshal.FinalReleaseComObject(iPropertyDescription);
                    iPropertyDescription = null;
                }
            }

            return propertyKey;
        }

        public void Dispose()
        {
            Dispose(true);
        }

        ~PropertySystem()
        {
            Dispose(false);
        }

        private void Dispose(bool disposing)
        {
            if (m_IPropertySystem != null)
            {
                if (!disposing)
                {
                    Debug.Fail("Failed to dispose PropertySystem");
                }

                Marshal.FinalReleaseComObject(m_IPropertySystem);
                m_IPropertySystem = null;
            }
            if (disposing)
            {
                GC.SuppressFinalize(this);
            }
        }

        /*
        Other members not wrapped
            GetPropertyDescriptionListFromString
            EnumeratePropertyDescriptions
            FormatForDisplay
            RegisterPropertySchema
            UnregisterPropertySchema
            RefreshPropertySchema
        */

    } // class PropertySystem

    /// <summary>
    /// Wrapper class for IPropertyDescription. Provides information about a particular Windows
    /// Property Store property.
    /// </summary>
    public class PropertyDescription
    {
        PROPERTYKEY m_propertyKey;
        string m_canonicalName;
        string m_displayName;
        PROPDESC_TYPE_FLAGS m_typeFlags;
        PROPDESC_VIEW_FLAGS m_viewFlags;
        ushort m_vt; // Variant type from which the managed type is derived.

#if !RAW_PROPERTY_STORE
        // These properties should be marked as innate or read-only but, at least in Windows 10,
        // they are not. If not RAW_PROPERTY_STORE compile-time option then we modify them.
        static HashSet<PROPERTYKEY> s_ForceInnate = new HashSet<PROPERTYKEY>(
            new PROPERTYKEY[]
            {
                new PROPERTYKEY("D5CDD502-2E9C-101B-9397-08002B2CF9AE", 26), // System.ContentType
                new PROPERTYKEY("D6942081-D53B-443D-AD47-5E059D9CD27A", 2), // System.Shell.SFGAOFlagsStrings
                new PROPERTYKEY("09329b74-40a3-4c68-bf07-af9a572f607c", 100), // System.IsFolder
                new PROPERTYKEY("14b81da1-0135-4d31-96d9-6cbfc9671a99", 18258), // System.DateImported
                new PROPERTYKEY("2e4b640d-5019-46d8-8881-55414cc5caa0", 100) // System.Media.DateEncoded
    });
#endif

        /// <summary>
        /// 
        /// </summary>
        /// <param name="iPropertyDescription"></param>
        /// <remarks>
        /// <para>In order to avoid having to make this class disposable, it retrieves all values
        /// from IPropertyDescription in the constructor. The caller should release iPropertyDescription.
        /// </para>
        /// </remarks>
        internal PropertyDescription(PropertyStoreInterop.IPropertyDescription iPropertyDescription)
        {
            Int32 hResult;

            // Get Property Key
            hResult = iPropertyDescription.GetPropertyKey(out m_propertyKey);
            if (hResult < 0)
            {
                Marshal.ThrowExceptionForHR(hResult);
            }

            // Get Canonical Name
            IntPtr pszName = IntPtr.Zero;
            try
            {
                hResult = iPropertyDescription.GetCanonicalName(out pszName);
                if (hResult >= 0 && pszName != IntPtr.Zero)
                {
                    m_canonicalName = Marshal.PtrToStringUni(pszName);
                }
                else
                {
                    Marshal.ThrowExceptionForHR(hResult);
                }
            }
            finally
            {
                if (pszName != IntPtr.Zero)
                {
                    Marshal.FreeCoTaskMem(pszName);
                    pszName = IntPtr.Zero;
                }
            }

            // Get Display Name
            pszName = IntPtr.Zero;
            try
            {
                hResult = iPropertyDescription.GetDisplayName(out pszName);
                if (hResult >= 0 && pszName != IntPtr.Zero)
                {
                    m_displayName = Marshal.PtrToStringUni(pszName);
                }
                else
                {
                    m_displayName = null;
                }
            }
            finally
            {
                if (pszName != IntPtr.Zero)
                {
                    Marshal.FreeCoTaskMem(pszName);
                    pszName = IntPtr.Zero;
                }
            }

            // Get Type
            hResult = iPropertyDescription.GetPropertyType(out m_vt);
            if (hResult < 0)
            {
                m_vt = 0;
                Debug.Fail("IPropertyDescription.GetPropertyType failed.");
            }

            // Get Type Flags
            hResult = iPropertyDescription.GetTypeFlags(PROPDESC_TYPE_FLAGS.PDTF_MASK_ALL, out m_typeFlags);
            if (hResult < 0)
            {
                m_typeFlags = 0;
                Debug.Fail("IPropertyDescription.GetTypeFlags failed.");
            }
#if !RAW_PROPERTY_STORE
            if (s_ForceInnate.Contains(m_propertyKey))
            {
                m_typeFlags |= PROPDESC_TYPE_FLAGS.PDTF_ISINNATE;
            }
#endif

            // Get View Flags
            hResult = iPropertyDescription.GetViewFlags(out m_viewFlags);
            if (hResult < 0)
            {
                m_typeFlags = 0;
                Debug.Fail("IPropertyDescription.GetViewFlags failed.");
            }

        }

        /// <summary>
        /// The <see cref="PROPERTYKEY"/> that identifies the property.
        /// </summary>
        public PROPERTYKEY PropertyKey
        {
            get { return m_propertyKey; }
        }

        /// <summary>
        /// The canonical name for the property.
        /// </summary>
        /// <remarks>
        /// A typical canonical name is "System.Image.HorizontalSize"
        /// </remarks>
        public string CanonicalName
        {
            get { return m_canonicalName; }
        }

        /// <summary>
        /// The display name for the property or null.
        /// </summary>
        /// <remarks>
        /// <para>A typical display name is "Width".</para>
        /// <para>Not all properties have display names. If a display name is not assigned, the value will be null.
        /// </para>
        /// </remarks>
        public string DisplayName
        {
            get { return m_displayName; }
        }

        /// <summary>
        /// The <see cref="PROPDESC_TYPE_FLAGS"/> for the property.
        /// </summary>
        /// <remarks>
        /// <see cref="PROPDESC_TYPE_FLAGS"/> is a bitmask and multiple flags may be set.
        /// </remarks>
        public PROPDESC_TYPE_FLAGS TypeFlags
        {
            get { return m_typeFlags; }
        }

        /// <summary>
        /// The <see cref="PROPDESC_VIEW_FLAGS"/> for the property.
        /// </summary>
        /// <remarks>
        /// <see cref="PROPDESC_VIEW_FLAGS"/> is a bitmask and multiple flags may be set.
        /// </remarks>
        public PROPDESC_VIEW_FLAGS ViewFlags
        {
            get { return m_viewFlags; }
        }

        /// <summary>
        /// Indicates whether this property value type is supported by the managed wrapper
        /// </summary>
        public bool ValueTypeIsSupported
        {
            get
            {
                switch (m_vt)
                {
                    case 0: // VT_EMPTY
                    case 1: // VT_NULL
                        return false;
                    case 2: // VT_I2
                    case 3: // VT_I4
                    case 4: // VT_R4
                    case 5: // VT_R8
                    case 6: // VT_CY
                    case 8: // VT_BSTR
                    case 10: // VT_ERROR
                    case 11: // VT_BOOL
                    case 14: // VT_DECIMAL
                    case 16: // VT_I1
                    case 17: // VT_UI1
                    case 18: // VT_UI2
                    case 19: // VT_UI4
                    case 20: // VT_I8
                    case 21: // VT_UI8
                    case 22: // VT_INT
                    case 23: // VT_UINT
                    case 25: // VT_HRESULT
                    case 30: // VT_LPSTR
                    case 31: // VT_LPWSTR
                    case 64: // VT_FILETIME
                    case 72: // VT_CLSID
                    case 0x1002: // VT_VECTOR|VT_I2
                    case 0x1003: // VT_VECTOR|VT_I4
                    case 0x1004: // VT_VECTOR|VT_R4
                    case 0x1005: // VT_VECTOR|VT_R8
                    case 0x1010: // VT_VECTOR|VT_I1
                    case 0x1011: // VT_VECTOR|VT_UI1
                    case 0x1012: // VT_VECTOR|VT_UI2
                    case 0x1013: // VT_VECTOR|VT_UI4
                    case 0x1014: // VT_VECTOR|VT_I8
                    case 0x1015: // VT_VECTOR|VT_UI8
                    case 0x1016: // VT_VECTOR|VT_INT
                    case 0x1017: // VT_VECTOR|VT_UINT
                    case 0x101E: // VT_VECTOR|VT_LPSTR
                    case 0x101F: // VT_VECTOR|VT_LPWSTR
                        return true;

                    case 66: // 0x42 VT_STREAM (Used by: System.ThumbnailStream on .mp3 format file, System.Contact.AccountPictureLarge, System.Contact.AccountPictureSmall)
                    case 71: // 0x47 VT_CF (CLIPDATA format, Used by: System.Thumbnail on .xls format file)
                    case 0x100C: // VT_VECTOR|VT_VARIANT (Used by: Unnamed property on .potx file)
                        return false;

                    default:
#if DEBUG
                        throw new NotImplementedException(string.Format("Unexpected PROPVARIANT type 0x{0:x4}.", m_vt));
#else
                        return false;
#endif
                }
            }
        }

        /// <summary>
        /// The managed type that is used to represent this property value.
        /// </summary>
        /// <remarks>
        /// Value is null if the property value type is not supported by the managed wrapper
        /// </remarks>
        public Type ValueType
        {
            get
            {
                switch (m_vt)
                {
                    case 0: // VT_EMPTY
                    case 1: // VT_NULL
                        return null;
                    case 2: // VT_I2
                        return typeof(Int16);
                    case 3: // VT_I4
                    case 22: // VT_INT
                        return typeof(Int32);
                    case 4: // VT_R4
                        return typeof(float);
                    case 5: // VT_R8
                        return typeof(double);
                    case 6: // VT_CY
                        return typeof(decimal);
                    case 8: // VT_BSTR
                    case 30: // VT_LPSTR
                    case 31: // VT_LPWSTR
                        return typeof(string);
                    case 10: // VT_ERROR
                        return typeof(UInt32);
                    case 11: // VT_BOOL
                        return typeof(bool);
                    case 14: // VT_DECIMAL
                        return typeof(decimal);
                    case 16: // VT_I1
                        return typeof(sbyte);
                    case 17: // VT_UI1
                        return typeof(byte);
                    case 18: // VT_UI2
                        return typeof(UInt16);
                    case 19: // VT_UI4
                    case 23: // VT_UINT
                    case 25: // VT_HRESULT
                        return typeof(UInt32);
                    case 20: // VT_I8
                        return typeof(Int64);
                    case 21: // VT_UI8
                        return typeof(UInt64);
                    case 64: // VT_FILETIME
                        return typeof(DateTime);
                    case 72: // VT_CLSID
                        return typeof(Guid);
                    case 0x1002: // VT_VECTOR|VT_I2
                        return typeof(Int16[]);
                    case 0x1003: // VT_VECTOR|VT_I4
                    case 0x1016: // VT_VECTOR|VT_INT
                        return typeof(Int32[]);
                    case 0x1004: // VT_VECTOR|VT_R4
                        return typeof(float[]);
                    case 0x1005: // VT_VECTOR|VT_R8
                        return typeof(double[]);
                    case 0x1010: // VT_VECTOR|VT_I1
                        return typeof(sbyte[]);
                    case 0x1011: // VT_VECTOR|VT_UI1
                        return typeof(byte[]);
                    case 0x1012: // VT_VECTOR|VT_UI2
                        return typeof(UInt16[]);
                    case 0x1013: // VT_VECTOR|VT_UI4
                    case 0x1017: // VT_VECTOR|VT_UINT
                        return typeof(UInt32[]);
                    case 0x1014: // VT_VECTOR|VT_I8
                        return typeof(Int64[]);
                    case 0x1015: // VT_VECTOR|VT_UI8
                        return typeof(UInt64[]);
                    case 0x101E: // VT_VECTOR|VT_LPSTR
                    case 0x101F: // VT_VECTOR|VT_LPWSTR
                        return typeof(String[]);

                    case 66: // 0x42 VT_STREAM (Used by: System.ThumbnailStream on .mp3 format file, System.Contact.AccountPictureLarge, System.Contact.AccountPictureSmall)
                    case 71: // 0x47 VT_CF (CLIPDATA format, Used by: System.Thumbnail on .xls format file)
                    case 0x100C: // VT_VECTOR|VT_VARIANT (Used by: Unnamed property on .potx file)
                        return null;

                    default:
#if DEBUG
                        throw new NotImplementedException(string.Format("Unexpected PROPVARIANT type 0x{0:x4}.", m_vt));
#else
                        return null;
#endif
                }
            }
        } // ValueType

        public bool Equals(PropertyDescription pd)
        {
            if (pd == null) return false;
            return m_propertyKey.Equals(pd.m_propertyKey);
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as PropertyDescription);
        }

        public override int GetHashCode()
        {
            return m_propertyKey.GetHashCode();
        }

        public override string ToString()
        {
            return m_canonicalName;
        }

        /*
        Other Members not implemented
            PropertyType
            EditInvitation
            DefaultColumnWidth
            DisplayType
            ColumnState
            GroupingRange
            RelativeDescriptionType
            RelativeDescription
            SortDescription
            SortDescriptionLabel
            AggregationType
            ConditionType
            EnumTypeList
            CoerceToCanonicalValue
            FormatForDisplay
            IsValueCanonical
        */
    } // class PropertyDescription

    [StructLayout (LayoutKind.Sequential, Pack = 4)]
    public struct PROPERTYKEY
    {
        public Guid fmtid;
        public UInt32 pid;

        public PROPERTYKEY(Guid guid, UInt32 propertyId)
        {
            fmtid = guid;
            pid = propertyId;
        }

        public PROPERTYKEY(string guid, UInt32 propertyId)
        {
            fmtid = new Guid(guid);
            pid = propertyId;
        }

        public bool Equals(PROPERTYKEY pk)
        {
            return fmtid.Equals(pk.fmtid) && pid == pk.pid;
        }

        public override bool Equals(object obj)
        {
            if (obj is PROPERTYKEY)
            {
                return Equals((PROPERTYKEY)obj);
            }
            return false;
        }

        public override int GetHashCode()
        {
            return fmtid.GetHashCode() ^ pid.GetHashCode();
        }

        public override string ToString()
        {
            return string.Concat("(", fmtid.ToString(), ",", pid.ToString(), ")");
        }
    } // struct PROPERTYKEY

    /// <summary>
    /// Declares the values of the <see cref="PropertyDescription.TypeFlags"/> property.
    /// </summary>
    /// <remarks>
    /// See <seealso cref="https://msdn.microsoft.com/en-us/library/windows/desktop/bb762527(v=vs.85).aspx"/> for definitions.
    /// </remarks>
    [Flags]
    public enum PROPDESC_TYPE_FLAGS : uint
    {
        PDTF_DEFAULT = 0x00000000,
        PDTF_MULTIPLEVALUES = 0x00000001,
        PDTF_ISINNATE = 0x00000002,
        PDTF_ISGROUP = 0x00000004,
        PDTF_CANGROUPBY = 0x00000008,
        PDTF_CANSTACKBY = 0x00000010,
        PDTF_ISTREEPROPERTY = 0x00000020,
        PDTF_INCLUDEINFULLTEXTQUERY = 0x00000040,
        PDTF_ISVIEWABLE = 0x00000080,
        PDTF_ISQUERYABLE = 0x00000100,
        PDTF_CANBEPURGED = 0x00000200,
        PDTF_SEARCHRAWVALUE = 0x00000400,
        PDTF_ISSYSTEMPROPERTY = 0x80000000,
        PDTF_MASK_ALL = 0x800007FF
    }

    /// <summary>
    /// Declares the values of the <see cref="PropertyDescription.ViewFlags"/> property.
    /// </summary>
    /// <remarks>
    /// See <seealso cref="https://msdn.microsoft.com/en-us/library/windows/desktop/bb762528(v=vs.85).aspx"/> for definitions.
    /// </remarks>
    [Flags]
    public enum PROPDESC_VIEW_FLAGS : uint
    {
        PDVF_DEFAULT = 0x00000000,
        PDVF_CENTERALIGN = 0x00000001,
        PDVF_RIGHTALIGN = 0x00000002,
        PDVF_BEGINNEWGROUP = 0x00000004,
        PDVF_FILLAREA = 0x00000008,
        PDVF_SORTDESCENDING = 0x00000010,
        PDVF_SHOWONLYIFPRESENT = 0x00000020,
        PDVF_SHOWBYDEFAULT = 0x00000040,
        PDVF_SHOWINPRIMARYLIST = 0x00000080,
        PDVF_SHOWINSECONDARYLIST = 0x00000100,
        PDVF_HIDELABEL = 0x00000200,
        PDVF_HIDDEN = 0x00000800,
        PDVF_CANWRAP = 0x00001000,
        PDVF_MASK_ALL = 0x00001BFF
    }

    internal static class PropertyStoreInterop
    {
        /*
        // The C++ Version
        interface IPropertyStore : IUnknown
        {
            HRESULT GetCount([out] DWORD *cProps);
            HRESULT GetAt([in] DWORD iProp, [out] PROPERTYKEY *pkey);
            HRESULT GetValue([in] REFPROPERTYKEY key, [out] PROPVARIANT *pv);
            HRESULT SetValue([in] REFPROPERTYKEY key, [in] REFPROPVARIANT propvar);
            HRESULT Commit();
        }
        */
        [ComImport, Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IPropertyStore
        {
            void GetCount([Out] out uint cProps);

            void GetAt([In] uint iProp, out PROPERTYKEY pkey);

            void GetValue([In] ref PROPERTYKEY key, [In] IntPtr pv);

            void SetValue([In] ref PROPERTYKEY key, [In] IntPtr pv);

            void Commit();
        }

        /*
        // The C++ Version
        MIDL_INTERFACE("c8e2d566-186e-4d49-bf41-6909ead56acc")
        interface IPropertyStoreCapabilities : public IUnknown
        {
            virtual HRESULT STDMETHODCALLTYPE IsPropertyWritable([in] REFPROPERTYKEY key);
        };
        */
        [ComImport, Guid("c8e2d566-186e-4d49-bf41-6909ead56acc"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IPropertyStoreCapabilities
        {
            [PreserveSig]
            Int32 IsPropertyWritable([In] ref PROPERTYKEY pkey);
        }

        /*
        // The C++ Version
        MIDL_INTERFACE("ca724e8a-c3e6-442b-88a4-6fb0db8035a3")
        IPropertySystem : public IUnknown
        {
        public:
            virtual HRESULT STDMETHODCALLTYPE GetPropertyDescription( 
                __RPC__in REFPROPERTYKEY propkey,
                __RPC__in REFIID riid,
                __RPC__deref_out_opt void **ppv) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetPropertyDescriptionByName( 
                __RPC__in_string LPCWSTR pszCanonicalName,
                __RPC__in REFIID riid,
                __RPC__deref_out_opt void **ppv) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetPropertyDescriptionListFromString( 
                __RPC__in_string LPCWSTR pszPropList,
                __RPC__in REFIID riid,
                __RPC__deref_out_opt void **ppv) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE EnumeratePropertyDescriptions( 
                PROPDESC_ENUMFILTER filterOn,
                __RPC__in REFIID riid,
                __RPC__deref_out_opt void **ppv) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE FormatForDisplay( 
                __RPC__in REFPROPERTYKEY key,
                __RPC__in REFPROPVARIANT propvar,
                PROPDESC_FORMAT_FLAGS pdff,
                __RPC__out_ecount_full_string(cchText) LPWSTR pszText,
                __RPC__in_range(0,0x8000) DWORD cchText) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE FormatForDisplayAlloc( 
                __RPC__in REFPROPERTYKEY key,
                __RPC__in REFPROPVARIANT propvar,
                PROPDESC_FORMAT_FLAGS pdff,
                __RPC__deref_out_opt_string LPWSTR *ppszDisplay) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE RegisterPropertySchema( 
                __RPC__in_string LPCWSTR pszPath) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE UnregisterPropertySchema( 
                __RPC__in_string LPCWSTR pszPath) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE RefreshPropertySchema( void) = 0;
        
        };
        */
        [ComImport, Guid("ca724e8a-c3e6-442b-88a4-6fb0db8035a3"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IPropertySystem
        {
            [PreserveSig]
            Int32 GetPropertyDescription([In] ref PROPERTYKEY propkey, [In] ref Guid riid, [Out] out IPropertyDescription rPropertyDescription);

            [PreserveSig]
            Int32 GetPropertyDescriptionByName([In][MarshalAs(UnmanagedType.LPWStr)] string pszCanonicalName, [In] ref Guid riid, [Out] out IPropertyDescription rPropertyDescription);

            // === All Other Methods Deferred Until Later! ===
        }

        /*
        // The C++ Version
        MIDL_INTERFACE("6f79d558-3e96-4549-a1d1-7d75d2288814")
        IPropertyDescription : public IUnknown
        {
        public:
            virtual HRESULT STDMETHODCALLTYPE GetPropertyKey( 
                __RPC__out PROPERTYKEY *pkey) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetCanonicalName( 
                __RPC__deref_out_opt_string LPWSTR *ppszName) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetPropertyType( 
                __RPC__out VARTYPE *pvartype) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetDisplayName( 
                __RPC__deref_out_opt_string LPWSTR *ppszName) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetEditInvitation( 
                __RPC__deref_out_opt_string LPWSTR *ppszInvite) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetTypeFlags( 
                PROPDESC_TYPE_FLAGS mask,
                __RPC__out PROPDESC_TYPE_FLAGS *ppdtFlags) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetViewFlags( 
                __RPC__out PROPDESC_VIEW_FLAGS *ppdvFlags) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetDefaultColumnWidth( 
                __RPC__out UINT *pcxChars) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetDisplayType( 
                __RPC__out PROPDESC_DISPLAYTYPE *pdisplaytype) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetColumnState( 
                __RPC__out SHCOLSTATEF *pcsFlags) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetGroupingRange( 
                __RPC__out PROPDESC_GROUPING_RANGE *pgr) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetRelativeDescriptionType( 
                __RPC__out PROPDESC_RELATIVEDESCRIPTION_TYPE *prdt) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetRelativeDescription( 
                __RPC__in REFPROPVARIANT propvar1,
                __RPC__in REFPROPVARIANT propvar2,
                __RPC__deref_out_opt_string LPWSTR *ppszDesc1,
                __RPC__deref_out_opt_string LPWSTR *ppszDesc2) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetSortDescription( 
                __RPC__out PROPDESC_SORTDESCRIPTION *psd) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetSortDescriptionLabel( 
                BOOL fDescending,
                __RPC__deref_out_opt_string LPWSTR *ppszDescription) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetAggregationType( 
                __RPC__out PROPDESC_AGGREGATION_TYPE *paggtype) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetConditionType( 
                __RPC__out PROPDESC_CONDITION_TYPE *pcontype,
                __RPC__out CONDITION_OPERATION *popDefault) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetEnumTypeList( 
                __RPC__in REFIID riid,
                __RPC__deref_out_opt void **ppv) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE CoerceToCanonicalValue( 
                _Inout_  PROPVARIANT *ppropvar) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE FormatForDisplay( 
                __RPC__in REFPROPVARIANT propvar,
                PROPDESC_FORMAT_FLAGS pdfFlags,
                __RPC__deref_out_opt_string LPWSTR *ppszDisplay) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE IsValueCanonical( 
                __RPC__in REFPROPVARIANT propvar) = 0;
        
        };
        */
        [ComImport, Guid("6f79d558-3e96-4549-a1d1-7d75d2288814"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IPropertyDescription
        {
            [PreserveSig]
            Int32 GetPropertyKey([Out] out PROPERTYKEY pkey);

            [PreserveSig]
            Int32 GetCanonicalName([Out] out IntPtr ppszName);

            [PreserveSig]
            Int32 GetPropertyType([Out] out ushort vartype);

            [PreserveSig]
            Int32 GetDisplayName([Out] out IntPtr ppszName);

            [PreserveSig]
            Int32 GetEditInvitation([Out] out IntPtr ppszInvite);

            [PreserveSig]
            Int32 GetTypeFlags([In] PROPDESC_TYPE_FLAGS mask, [Out] out PROPDESC_TYPE_FLAGS ppdtFlags);

            [PreserveSig]
            Int32 GetViewFlags([Out] out PROPDESC_VIEW_FLAGS ppdtFlags);

            // === All Other Methods Deferred Until Later! ===
        }

        /*
        // C++ version
        typedef struct PROPVARIANT {
            VARTYPE vt;
            WORD    wReserved1;
            WORD    wReserved2;
            WORD    wReserved3;
            union {
                // Various types of up to 8 bytes
            }
        } PROPVARIANT;
        */
        [StructLayout(LayoutKind.Explicit)]
        internal struct PROPVARIANT
        {
            [FieldOffset(0)]
            public ushort vt;
            [FieldOffset(2)]
            public ushort wReserved1;
            [FieldOffset(4)]
            public ushort wReserved2;
            [FieldOffset(6)]
            public ushort wReserved3;
            [FieldOffset(8)]
            public Int32 data01;
            [FieldOffset(12)]
            public Int32 data02;

            // IntPtr (for strings and the like)
            [FieldOffset(8)]
            public IntPtr dataIntPtr;

            // For FileTime and Int64
            [FieldOffset(8)]
            public long dataInt64;

            // Vector-style arrays (for VT_VECTOR|VT_LPWSTR and such)
            [FieldOffset(8)]
            public uint cElems;
            [FieldOffset(12)]
            public IntPtr pElems32;
            [FieldOffset(16)]
            public IntPtr pElems64;

            public IntPtr pElems
            {
                get { return (IntPtr.Size == 4) ? pElems32 : pElems64; }
            }
        }

        [Flags]
        public enum GETPROPERTYSTOREFLAGS : uint
        {
            // If no flags are specified (GPS_DEFAULT), a read-only property store is returned that includes properties for the file or item.
            // In the case that the shell item is a file, the property store contains:
            //     1. properties about the file from the file system
            //     2. properties from the file itself provided by the file's property handler, unless that file is offline,
            //     see GPS_OPENSLOWITEM
            //     3. if requested by the file's property handler and supported by the file system, properties stored in the
            //     alternate property store.
            //
            // Non-file shell items should return a similar read-only store
            //
            // Specifying other GPS_ flags modifies the store that is returned
            GPS_DEFAULT = 0x00000000,
            GPS_HANDLERPROPERTIESONLY = 0x00000001,   // only include properties directly from the file's property handler
            GPS_READWRITE = 0x00000002,   // Writable stores will only include handler properties
            GPS_TEMPORARY = 0x00000004,   // A read/write store that only holds properties for the lifetime of the IShellItem object
            GPS_FASTPROPERTIESONLY = 0x00000008,   // do not include any properties from the file's property handler (because the file's property handler will hit the disk)
            GPS_OPENSLOWITEM = 0x00000010,   // include properties from a file's property handler, even if it means retrieving the file from offline storage.
            GPS_DELAYCREATION = 0x00000020,   // delay the creation of the file's property handler until those properties are read, written, or enumerated
            GPS_BESTEFFORT = 0x00000040,   // For readonly stores, succeed and return all available properties, even if one or more sources of properties fails. Not valid with GPS_READWRITE.
            GPS_NO_OPLOCK = 0x00000080,   // some data sources protect the read property store with an oplock, this disables that
            GPS_MASK_VALID = 0x000000FF,
        }

        [DllImport("shell32.dll", SetLastError = true, CallingConvention = CallingConvention.StdCall, PreserveSig = false)]
        public static extern void SHGetPropertyStoreFromParsingName(
                [In][MarshalAs(UnmanagedType.LPWStr)] string pszPath,
                [In] IntPtr zeroWorks,
                [In] GETPROPERTYSTOREFLAGS flags,
                [In] ref Guid iIdPropStore,
                [Out] out IPropertyStore propertyStore);

        [DllImport(@"ole32.dll", SetLastError = true, CallingConvention = CallingConvention.StdCall, PreserveSig = false)]
        public static extern void PropVariantInit([In] IntPtr pvarg);

        [DllImport(@"ole32.dll", SetLastError = true, CallingConvention = CallingConvention.StdCall, PreserveSig = false)]
        public static extern void PropVariantClear([In] IntPtr pvarg);

        [DllImport("propsys.dll", SetLastError = true, CallingConvention = CallingConvention.StdCall, PreserveSig = false)]
        public static extern void PSGetPropertySystem([In] ref Guid iIdPropertySystem, [Out] out IPropertySystem propertySystem);

        // Converts a string to a PropVariant with type LPWSTR instead of BSTR
        // The resulting variant must be cleared using PropVariantClear and freed using Marshal.FreeCoTaskMem
        public static IntPtr PropVariantFromString(string value)
        {
            IntPtr pstr = IntPtr.Zero;
            IntPtr pv = IntPtr.Zero;
            try
            {
                // In managed code, new automatically zeros the contents.
                PROPVARIANT propvariant = new PROPVARIANT();

                // Allocate the string
                pstr = Marshal.StringToCoTaskMemUni(value);

                // Allocate the PropVariant
                pv = Marshal.AllocCoTaskMem(16);

                // Transfer ownership of the string
                propvariant.vt = 31; // VT_LPWSTR - not documented but this is to be allocated using CoTaskMemAlloc.
                propvariant.dataIntPtr = pstr;
                Marshal.StructureToPtr(propvariant, pv, false);
                pstr = IntPtr.Zero;

                // Transfer ownership to the result
                IntPtr result = pv;
                pv = IntPtr.Zero;
                return result;
            }
            finally
            {
                if (pstr != IntPtr.Zero)
                {
                    Marshal.FreeCoTaskMem(pstr);
                    pstr = IntPtr.Zero;
                }
                if (pv != IntPtr.Zero)
                {
                    try
                    {
                        PropertyStoreInterop.PropVariantClear(pv);
                    }
                    catch
                    {
                        Debug.Fail("VariantClear failure");
                    }
                    Marshal.FreeCoTaskMem(pv);
                    pv = IntPtr.Zero;
                }
            }
        }

        // Converts an object to a PropVariant including special handling for strings
        // The resulting variant must be cleared using PropVariantClear and freed using Marshal.FreeCoTaskMem
        public static IntPtr PropVariantFromObject(object value)
        {
            {
                string strValue = value as string;
                if (strValue != null)
                {
                    return PropVariantFromString(strValue);
                }
            }

            {
                IntPtr pv = IntPtr.Zero;
                try
                {
                    pv = Marshal.AllocCoTaskMem(16);

                    if (value is DateTime)
                    {
                        DateTime dtValue = (DateTime)value;
                        PROPVARIANT v = new PROPVARIANT();
                        v.vt = 64; // VT_FILETIME
                        v.dataInt64 = dtValue.ToFileTimeUtc();
                        Marshal.StructureToPtr(v, pv, false);
                    }
                    else
                    {
                        Marshal.GetNativeVariantForObject(value, pv);
                    }

                    IntPtr result = pv;
                    pv = IntPtr.Zero;
                    return result;
                }
                finally
                {
                    if (pv != IntPtr.Zero)
                    {
                        try
                        {
                            PropertyStoreInterop.PropVariantClear(pv);
                        }
                        catch
                        {
                            Debug.Fail("VariantClear failure");
                        }
                        Marshal.FreeCoTaskMem(pv);
                        pv = IntPtr.Zero;
                    }
                }
            }
        } // Method PropVariantFromObject

        // Reference: https://msdn.microsoft.com/en-us/library/windows/desktop/aa380072(v=vs.85).aspx
        public static object PropVariantToObject(IntPtr pv)
        {
            // Copy to structure
            PROPVARIANT v = (PROPVARIANT)Marshal.PtrToStructure(pv, typeof(PROPVARIANT));

            object value = null;
            switch (v.vt)
            {
                case 0: // VT_EMPTY
                case 1: // VT_NULL
                case 2: // VT_I2
                case 3: // VT_I4
                case 4: // VT_R4
                case 5: // VT_R8
                case 6: // VT_CY
                //case 7: // VT_DATE
                case 8: // VT_BSTR
                case 10: // VT_ERROR
                case 11: // VT_BOOL
                case 14: // VT_DECIMAL
                case 16: // VT_I1
                case 17: // VT_UI1
                case 18: // VT_UI2
                case 19: // VT_UI4
                case 20: // VT_I8
                case 21: // VT_UI8
                case 22: // VT_INT
                case 23: // VT_UINT
                case 24: // VT_VOID
                case 25: // VT_HRESULT
                    value = Marshal.GetObjectForNativeVariant(pv);
                    break;

                case 30: // VT_LPSTR
                    value = Marshal.PtrToStringAnsi(v.dataIntPtr);
                    break;

                case 31: // VT_LPWSTR
                    value = Marshal.PtrToStringUni(v.dataIntPtr);
                    break;

                case 64: // VT_FILETIME
                    value = DateTime.FromFileTimeUtc(v.dataInt64);
                    break;

                case 66: // 0x42 VT_STREAM (Used by: System.ThumbnailStream on .mp3 format file, System.Contact.AccountPictureLarge, System.Contact.AccountPictureSmall)
                    throw new NotImplementedException("Conversion of PROPVARIANT VT_STREAM to managed type not yet supported.");

                case 71: // 0x47 VT_CF (CLIPDATA format, Used by: System.Thumbnail on .xls format file)
                    throw new NotImplementedException("Conversion of PROPVARIANT VT_CF to managed type not yet supported.");

                case 72: // VT_CLSID
                    {
                        byte[] bytes = new byte[16];
                        Marshal.Copy(v.dataIntPtr, bytes, 0, 16);
                        value = new Guid(bytes);
                    }
                    break;

                case 0x1002: // VT_VECTOR|VT_I2
                    {
                        Int16[] a = new Int16[v.cElems];
                        Marshal.Copy(v.pElems, a, 0, (int)v.cElems);
                        value = a;
                    }
                    break;

                case 0x1003: // VT_VECTOR|VT_I4
                case 0x1016: // VT_VECTOR|VT_INT
                    {
                        Int32[] a = new Int32[v.cElems];
                        Marshal.Copy(v.pElems, a, 0, (int)v.cElems);
                        value = a;
                    }
                    break;

                case 0x1004: // VT_VECTOR|VT_R4
                    {
                        float[] a = new float[v.cElems];
                        Marshal.Copy(v.pElems, a, 0, (int)v.cElems);
                        value = a;
                    }
                    break;

                case 0x1010: // VT_VECTOR|VT_I1
                    {
                        byte[] a = new byte[v.cElems];
                        Marshal.Copy(v.pElems, a, 0, (int)v.cElems);
                        SByte[] b = new SByte[v.cElems];
                        Buffer.BlockCopy(a, 0, b, 0, (int)v.cElems);
                        value = b;
                    }
                    break;

                case 0x1011: // VT_VECTOR|VT_UI1
                    {
                        byte[] a = new byte[v.cElems];
                        Marshal.Copy(v.pElems, a, 0, (int)v.cElems);
                        value = a;
                    }
                    break;

                case 0x1012: // VT_VECTOR|VT_UI2
                    {
                        Int16[] a = new Int16[v.cElems];
                        Marshal.Copy(v.pElems, (Int16[])a, 0, (int)v.cElems);
                        UInt16[] b = new UInt16[v.cElems];
                        Buffer.BlockCopy(a, 0, b, 0, (int)v.cElems * 2);
                        value = b;
                    }
                    break;

                case 0x1013: // VT_VECTOR|VT_UI4
                case 0x1017: // VT_VECTOR|VT_UINT
                    {
                        Int32[] a = new Int32[v.cElems];
                        Marshal.Copy(v.pElems, a, 0, (int)v.cElems);
                        UInt32[] b = new UInt32[v.cElems];
                        Buffer.BlockCopy(a, 0, b, 0, (int)v.cElems * 4);
                        value = b;
                    }
                    break;

                case 0x1014: // VT_VECTOR|VT_I8
                    {
                        Int64[] a = new Int64[v.cElems];
                        Marshal.Copy(v.pElems, a, 0, (int)v.cElems);
                        value = a;
                    }
                    break;

                case 0x1015: // VT_VECTOR|VT_UI8
                    {
                        Int64[] a = new Int64[v.cElems];
                        Marshal.Copy(v.pElems, a, 0, (int)v.cElems);
                        UInt64[] b = new UInt64[v.cElems];
                        Buffer.BlockCopy(a, 0, b, 0, (int)v.cElems * 8);
                        value = b;
                    }
                    break;

                case 0x1005: // VT_VECTOR|VT_R8
                    {
                        double[] doubles = new double[v.cElems];
                        Marshal.Copy(v.pElems, doubles, 0, (int)v.cElems);
                        value = doubles;
                    }
                    break;

                case 0x100C: // VT_VECTOR|VT_VARIANT (Used by: Unnamed property on .potx file)
                    throw new NotImplementedException("Conversion of PROPVARIANT VT_VECTOR|VT_VARIANT to managed type not yet supported.");

                case 0x101E: // VT_VECTOR|VT_LPSTR (Used by unnamed properties on .potx and .dotx files)
                    {
                        string[] strings = new string[v.cElems];
                        for (int i = 0; i < v.cElems; ++i)
                        {
                            IntPtr strPtr = Marshal.ReadIntPtr(v.pElems + i * IntPtr.Size);
                            strings[i] = Marshal.PtrToStringAnsi(strPtr);
                        }
                        value = strings;
                    }
                    break;

                case 0x101f: // VT_VECTOR|VT_LPWSTR
                    {
                        string[] strings = new string[v.cElems];
                        for (int i = 0; i < v.cElems; ++i)
                        {
                            IntPtr strPtr = Marshal.ReadIntPtr(v.pElems + i * IntPtr.Size);
                            strings[i] = Marshal.PtrToStringUni(strPtr);
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
                        value = String.Format("(Supported type 0x{0:x4}): {1}", v.vt, value.ToString());
                        throw new NotImplementedException(string.Format("Conversion of PROPVARIANT type 0x{0:x4} is not yet supported but Marshal.GetObjectForNativeVariant seems to work. value='{0}'", v.vt, value));
                    }
                    catch
                    {
                        // Do nothing
                    }
#endif
                    throw new NotImplementedException(string.Format("Conversion of PROPVARIANT type 0x{0:x4} is not yet supported.", v.vt));
            }

            return value;
        } // Method PropVariantToObject

    } // class PropertyStoreInterop

} // namespace WinShell
