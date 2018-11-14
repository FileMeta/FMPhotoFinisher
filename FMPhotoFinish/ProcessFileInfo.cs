using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace FMPhotoFinish
{
    class ProcessFileInfo
    {
        public ProcessFileInfo(FileInfo fi)
        {
            Filepath = fi.FullName;
            Size = fi.Length;
            OriginalFilepath = fi.FullName;
            OriginalDateCreated = fi.CreationTimeUtc;
            OriginalDateModified = fi.LastWriteTimeUtc;

        }

        public string Filepath { get; set; }
        public long Size { get; private set; }
        public string OriginalFilepath { get; private set; }
        public DateTime OriginalDateCreated { get; private set; }
        public DateTime OriginalDateModified { get; private set; }
    }
}
