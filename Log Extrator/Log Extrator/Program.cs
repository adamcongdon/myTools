// This original program was created by Ben Creamer
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Compression;
using System.IO;
using System.Windows;


namespace Log_Extrator
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Write("Enter Log Directory: ");
            var logDir = Console.ReadLine();
            //Extract .zip files
            var zipFiles = Directory.GetFiles(logDir, "*.zip", SearchOption.AllDirectories);
            foreach (var zipFile in zipFiles)
            {
                try
                {
                    var outDir = Path.GetDirectoryName(zipFile) + "\\" + Path.GetFileNameWithoutExtension(zipFile);
                    ZipFile.ExtractToDirectory(zipFile, outDir);
                    Console.WriteLine("Decompressed: {0}", Path.GetFileName(zipFile));
                    File.Delete(Path.GetFullPath(zipFile));
                }

                catch
                {
                    Console.WriteLine("Header error on {0}", zipFile);
                }
            }
                //Extract .gz files
                DirectoryInfo directorySelected = new DirectoryInfo(logDir);
                foreach (FileInfo gzFile in directorySelected.GetFiles("*.gz", SearchOption.AllDirectories))
                {

                    using (FileStream originalFileStream = gzFile.OpenRead())
                    {
                        string currentFileName = gzFile.FullName;
                        string simpleFileName = gzFile.Name;
                        string newFileName = currentFileName.Remove(currentFileName.Length - (gzFile.Extension.Length + 4)) + "\\" + simpleFileName.Remove(simpleFileName.Length - (gzFile.Extension.Length));
                        Directory.CreateDirectory(Convert.ToString(currentFileName.Remove(currentFileName.Length - (gzFile.Extension.Length + 4))));
                        using (FileStream decompressedFileStream = File.Create(newFileName))
                        {
                            using (GZipStream decompressionStream = new GZipStream(originalFileStream, CompressionMode.Decompress))
                            {
                                decompressionStream.CopyTo(decompressedFileStream);
                                Console.WriteLine("Decompressed: {0}", gzFile.Name);

                            }
                        File.Delete(gzFile.FullName);
                    }
                    }
                }
        }
    }
}
