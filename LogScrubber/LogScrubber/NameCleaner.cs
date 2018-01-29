using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace LogScrubber
{
    class NameCleaner
    {
        public static void Shorten(List<string> fileList)
        {
            foreach (string file in fileList)
            {
                const string remStr = "_CLN_CLN";
                if (file.Contains(remStr))
                {
                    string src = Path.GetFullPath(file);
                    string name = Path.GetFileName(file);
                    string newName = name.Replace(remStr, "");
                    string finName = newName + "-SCRBD.log";
                    string tarDir = Path.GetDirectoryName(file);
                    File.Move(src, tarDir + "\\" + finName);
                }
            }

        }
    }
}
