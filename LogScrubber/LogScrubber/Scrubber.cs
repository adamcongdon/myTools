using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;

namespace LogScrubber
{
    class Scrubber
    {
        public static void Scrub(List<string> scrubList)
        {
            foreach(string file in scrubList)
            {
                using (StreamReader sr = new StreamReader(file))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        string[] checkList;
                        Scrubber.Params(out checkList);
                        
                        foreach(string item in checkList)
                        {
                            if (line.Contains(item))
                            {
                                Console.WriteLine(line);
                            }
                        }
                    }

                }
            }
        }
        public static void ScrubIPs(List<string> scrubList)
        {
            Regex ip = new Regex(@"\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b");
            foreach (string file in scrubList)
            {
                using (StreamReader sr = new StreamReader(file))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        Match match = Regex.Match(line, @"\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b");

                        if(match.Success)
                        {
                            Console.WriteLine("Before: " + "\n" + line + "\n\n" + "AFTER" + "\n");
                            string newLine = line.Replace(match.ToString(), "ScrubbedIpWasHere");
                            
                            Console.WriteLine((newLine + "\n\n"));
                        }
                    }

                }
            }
        }
        public static string[] Params(out string[] phrases)
        {
            string[] pList =
            {
                "test1",
                "test2",
                "MachineName",
                //"IPAddresses",
                

            };

            phrases = pList;
            return phrases;
        }
    }
}
