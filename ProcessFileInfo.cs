using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace FMPhotoFinisher
{
    class ProcessFileInfo
    {
        public ProcessFileInfo(string filepath, long size)
        {
            Filepath = filepath;
            Size = size;
            OriginalFilename = Path.GetFileName(filepath);
        }

        public string Filepath { get; set; }
        public long Size { get; private set; }
        public string OriginalFilename { get; private set; }
    }
}
