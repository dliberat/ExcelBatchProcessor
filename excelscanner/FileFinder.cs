using System;
using System.Collections.Generic;
using System.IO;

namespace excelscanner
{
    public class FileFinder
    {
        private List<FileInfo> filePaths;
        private string[] InputPath;

        public FileFinder(string[] InputPath)
        {
            filePaths = new List<FileInfo>();
            this.InputPath = InputPath;
        }

        public List<FileInfo> LoadFiles()
        {
            foreach (string path in InputPath)
            {
                if (File.Exists(path))
                {
                    // This path is a file
                    LoadFile(path);
                }
                else if (Directory.Exists(path))
                {
                    // This path is a directory
                    LoadDirectory(path);
                }
                else
                {
                    Console.WriteLine("{0} is not a valid file or directory.", path);
                }
            }

            return filePaths;
        }

        // Process all files in the directory passed in, recurse on any directories 
        // that are found, and process the files they contain.
        private void LoadDirectory(string targetDirectory)
        {
            // Process the list of files found in the directory.
            string[] fileEntries = Directory.GetFiles(targetDirectory);
            foreach (string fileName in fileEntries)
                LoadFile(fileName);

            // Recurse into subdirectories of this directory.
            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            foreach (string subdirectory in subdirectoryEntries)
                LoadDirectory(subdirectory);
        }

        private void LoadFile(string path)
        {
            filePaths.Add(new FileInfo(path));
        }

    }
}
