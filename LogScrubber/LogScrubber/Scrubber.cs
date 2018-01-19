using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;

/*
 * Scrub musts:
 * user names
 * domain names
 * IP (3rd/4th octect)
 * Computer names
 * Can have IP or Hostname, not both.
 */

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
            List<string> ipList = new List<string>();
            foreach (string file in scrubList)
            {
                using (StreamReader sr = new StreamReader(file))
                    {
                        string outPath = Path.GetFullPath(file);
                        string outFile = @outPath + "__SCRUBBED.log";
                        using (StreamWriter sw = new StreamWriter(outFile))
                        { 
                        string line;
                            while ((line = sr.ReadLine()) != null)
                            {
                                Match match = Regex.Match(line, @"\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}");

                                const string ip1 = "192";
                                const string ip2 = "10.";
                                const string ip3 = "172";
                                const string ver = "9.5";
                                const string ver1 = "9.0";
                                const string ver2 = "8.0";
                                const string lb = "127.0.0.1"; 

                                if (line.Contains("startTime"))
                                {
                                    sw.WriteLine(line);
                                    continue;
                                }
                                if (line.Contains(ver))
                                {
                                    sw.WriteLine(line);
                                    continue;
                                }
                                if (line.Contains(ver1))
                                {
                                    sw.WriteLine(line);
                                    continue;
                                }
                                if (line.Contains(ver2))
                                {
                                    sw.WriteLine(line);
                                    continue;
                                }
                                if (line.Contains(lb))
                                {
                                    sw.WriteLine(line);
                                    continue;
                                }

                                if (match.Success)
                                {
                                    string matchString = match.ToString();
                                Match trimMatch = Regex.Match(matchString, @"\d{1,3}\.\d{1,3}\.");
                                    string replacer = trimMatch.ToString() + "x.x";
                                    if(matchString == lb)
                                    {
                                        sw.WriteLine(line);
                                        continue;
                                    }
                                    else
                                    {
                                        string newLine = line.Replace(matchString, replacer);
                                        sw.WriteLine(newLine);
                                        ipList.Add(matchString);
                                    }
                                }
                                else
                                {
                                    sw.WriteLine(line);
                                }
                            }
                        }
                }
                File.Delete(file);
            }
            Console.WriteLine("List of all found IPs: \n");
            foreach (string addr in ipList.Distinct())
            {
                Console.WriteLine(addr);
            }
            
            Console.WriteLine("PRESS ANY KEY TO CONTINUE");
            Console.ReadKey();
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
