using System;
using WinShell;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FMPhotoFinisher
{
    /// <summary>
    /// Performs operations on a media file such as metadata changes, rotation, recoding, etc.
    /// </summary>
    class MediaFile : IDisposable
    {
        string m_filepath;
        Dictionary<PROPERTYKEY, object> m_properties = new Dictionary<PROPERTYKEY, object>();

        public MediaFile(string filepath)
        {
            m_filepath = filepath;
            Orientation = 1; // Defaults to normal/vertical

            // Load metadata properties
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
        }

        public int Orientation { get; private set; }

        public void RotateToVertical()
        {
            JpegRotator.RotateToVertical(m_filepath);
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

        #region Static Members

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
