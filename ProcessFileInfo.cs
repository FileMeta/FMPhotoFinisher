using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FMPhotoFinisher
{
    class ProcessFileInfo
    {
        public ProcessFileInfo(string filename, long size)
        {
            Filename = filename;
            Size = size;
        }

        public string Filename { get; private set; }
        public long Size { get; private set; }
    }
}
