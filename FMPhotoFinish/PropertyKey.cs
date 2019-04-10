/*
---
name: PropKey.cs
description: Managed Code Interoperability for PROPKEY
url: https://github.com/FileMeta/WinShellPropertyStore/raw/master/PropertyKey.cs
version: 1.0
keywords: CodeBit
dateModified: 2019-04-09
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

using System;
using System.Runtime.InteropServices;

namespace Interop
{

    /// <summary>
    /// A property is identified by a PropertySetId or FormatId (GUID) and a PropertyId (UInt32).
    /// </summary>
    /// <remarks>
    /// <para>Property IDs are used at least two places in the Windows infrastructure. In the
    /// Windows Property System they are used to identify properties that may be attached to
    /// files or other items. The same set of PropertyIds are used by Windows Search when indexing
    /// files and other items. In OLE DB, PropertyIds are used to identify properties that may
    /// be attached to databases, queries, rowsets, or rows.
    /// </para>
    /// </remarks>
    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct PropertyKey : IComparable<PropertyKey>
    {
        /// <summary>
        /// The <see cref="Guid"/> id of the property set. Equivalent to fmtid or PropertySetId
        /// </summary>
        public Guid PropSetId;

        /// <summary>
        /// The id of the property within the property set.
        /// </summary>
        public UInt32 PropertyId;

        public PropertyKey(Guid propSetId, UInt32 propertyId)
        {
            PropSetId = propSetId;
            PropertyId = propertyId;
        }

        public PropertyKey(string propSetId, UInt32 propertyId)
        {
            PropSetId = new Guid(propSetId);
            PropertyId = propertyId;
        }

        public int CompareTo(PropertyKey other)
        {
            int diff = PropSetId.CompareTo(other.PropSetId);
            if (diff != 0) return diff;
            return PropertyId.CompareTo(other.PropertyId);
        }

        public bool Equals(PropertyKey pk)
        {
            return PropSetId.Equals(pk.PropSetId) && PropertyId == pk.PropertyId;
        }

        public override bool Equals(object obj)
        {
            if (obj is PropertyKey)
            {
                return Equals((PropertyKey)obj);
            }
            return false;
        }

        public override int GetHashCode()
        {
            return PropSetId.GetHashCode() ^ PropertyId.GetHashCode();
        }

        public override string ToString()
        {
            return string.Concat("(", PropSetId.ToString(), ",", PropertyId.ToString(), ")");
        }
    } // struct PropertyKey

}
