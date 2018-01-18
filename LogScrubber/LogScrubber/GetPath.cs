using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LogScrubber
{
    class GetPath
    {
        public static string Path(out string path)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Enter path with logs to scrub: ");
            Console.ForegroundColor = ConsoleColor.White;
            path = Console.ReadLine();
            return path;
        }
    }
}
