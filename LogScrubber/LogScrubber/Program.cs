using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LogScrubber
{
    class Program
    {
        static void Main(string[] args)
        {
            string setPath = null;
            
            FileLister.List(GetPath.Path(out setPath), out List<string> myFiles);

            //Scrubber.Scrub(myFiles);
            Scrubber.ScrubIPs(myFiles);

            // Below is Console Feedback for testing
            //Console.ForegroundColor = ConsoleColor.Magenta;
            //Console.WriteLine(setPath);
            //foreach (string file in myFiles)
            //{
            //    Console.WriteLine((file));
            //}
            //Console.WriteLine(myFiles.Count);
            //Console.ForegroundColor = ConsoleColor.White;
        }
    }
}
