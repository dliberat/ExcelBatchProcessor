using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelscanner
{
    public class FileProcessor
    {
        public void Process(string[] args)
        {
            foreach (string path in args)
            {
                if (File.Exists(path))
                {
                    // This path is a file
                    ProcessFile(path);
                }
                else if (Directory.Exists(path))
                {
                    // This path is a directory
                    ProcessDirectory(path);
                }
                else
                {
                    Console.WriteLine("{0} is not a valid file or directory.", path);
                }
            }
        }


        // Process all files in the directory passed in, recurse on any directories 
        // that are found, and process the files they contain.
        public void ProcessDirectory(string targetDirectory)
        {
            // Process the list of files found in the directory.
            string[] fileEntries = Directory.GetFiles(targetDirectory);
            foreach (string fileName in fileEntries)
                ProcessFile(fileName);

            // Recurse into subdirectories of this directory.
            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            foreach (string subdirectory in subdirectoryEntries)
                ProcessDirectory(subdirectory);
        }

        // Insert logic for processing found files here.
        public static void ProcessFile(string path)
        {
            FileInfo fi = new FileInfo(path);
            if (!fi.Extension.Contains("xl"))
            {
                Console.WriteLine("Skipping file '{0}'.", fi.Name);
                return;
            }

            ExcelScanner scanner = new ExcelScanner(fi);
            scanner.DeleteNonKRJPHeaders();

            string msg = $"Processed file '{path}'.";
            if (scanner.Modified) msg += " FILE MODIFIED";

            Console.WriteLine("Processed file '{0}'.", path);
        }
    }
}
