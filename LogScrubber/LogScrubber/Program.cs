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
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Enter path with logs to scrub: ");
            Console.ForegroundColor = ConsoleColor.White;
            setPath = Console.ReadLine();

            int x = 1;
            while (x < 4)
            {
                FileLister.List(setPath, out List<string> myFiles);
                Console.ForegroundColor = ConsoleColor.Magenta;
                Console.WriteLine("\nScrub pass # {0}\n", x);
                Console.ForegroundColor = ConsoleColor.White;

                //Scrubber.Scrub(myFiles);
                Scrubber.ScrubIPs(myFiles);
                x++;
            }

            NameCleaner.Shorten(FileLister.List(setPath, out List<string> iList));

            Console.WriteLine("PRESS ANY KEY TO CONTINUE");
            Console.ReadKey();
        }
    }
}
