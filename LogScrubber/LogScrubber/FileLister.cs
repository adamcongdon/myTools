using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace LogScrubber
{
    class FileLister
    {
        public static List<string> List(string path, out List<string> fileList)
        {
            List<string> newList = new List<string>(Directory.GetFiles(path, "*.log", SearchOption.AllDirectories));

            fileList = newList;
            return fileList;
        }
        
    }
}
