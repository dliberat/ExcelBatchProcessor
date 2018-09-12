using NLog;
using System.Collections.Generic;
using System.IO;
using System.IO.Abstractions;

namespace ExcelBatchProcessor.Files
{
    /// <summary>
    /// Finds files recursively in an array of provided locations.
    /// </summary>
    public class FileFinder : IFileFinder
    {
        readonly IFileSystem FileSystem;
        private static Logger logger = LogManager.GetCurrentClassLogger();
        private List<FileInfo> filePaths;
        private string[] InputPath;

        /// <summary>
        /// Full constructor used for testing purposes. Consumers of the <c>FileFinder</c>
        /// class should use the overloaded constructor instead.
        /// </summary>
        /// <param name="FileSystem"></param>
        /// <param name="InputPath"></param>
        public FileFinder(IFileSystem FileSystem, string[] InputPath)
        {
            this.FileSystem = FileSystem;
            filePaths = new List<FileInfo>();
            this.InputPath = InputPath;
        }
        /// <summary>
        /// Primary constructor for the <c>FileFinder</c> class.
        /// </summary>
        /// <param name="InputPath">Array of strings indicating target files or directories to search for files.</param>
        /// <example>
        /// The <c>FileFinder</c> can be provided an array of specific filenames, directories to search recursively,
        /// or a mix of both.
        /// <code>
        /// FileFinder ff = new FileFinder(new string[] { "c:\my_excel_file.xlsx" });
        /// var found = ff.Find(); // returns 1 file
        /// </code>
        /// <code>
        /// FileFinder ff = new FileFinder(new string[] { "c:\my_excel_files\" });
        /// var found = ff.Find(); // returns everything inside of c:\my_excel_files
        /// </code>
        /// <code>
        /// FileFinder ff = new FileFinder(new string[] { "c:\my_excel_files\", "d:\bonusfile.xlsx" });
        /// var found = ff.Find(); // returns everything in c:\my_excel_files plus bonusfile.xslx
        /// </code>
        /// </example>
        public FileFinder(string[] InputPath) : this(FileSystem: new FileSystem(), InputPath: InputPath)
        {
        }
        /// <summary>
        /// Traverses the <c>InputPath</c>s provided in the class constructor and collate a list
        /// of all the files found therein.
        /// </summary>
        /// <returns>All the valid files found recursively in the directories provided.</returns>
        public IEnumerable<FileInfo> Find()
        {
            foreach (string path in InputPath)
            {
                if (FileSystem.File.Exists(path))
                {
                    // This path is a file
                    LoadFile(path);
                }
                else if (FileSystem.Directory.Exists(path))
                {
                    // This path is a directory
                    LoadDirectory(path);
                }
                else
                {
                    logger.Warn("{path} is not a valid file or directory.", path);
                }
            }

            logger.Info("Found {count} files recursively in {path}", filePaths.Count, InputPath);
            return filePaths;
        }

        /// <summary>
        /// Process all files in the directory passed in, recurse on any directories 
        /// that are found, and process the files they contain.
        /// </summary>
        /// <param name="targetDirectory">Directory to be searched.</param>
        protected void LoadDirectory(string targetDirectory)
        {
            // Process the list of files found in the directory.
            string[] fileEntries = FileSystem.Directory.GetFiles(targetDirectory);
            foreach (string fileName in fileEntries)
                LoadFile(fileName);

            // Recurse into subdirectories of this directory.
            string[] subdirectoryEntries = FileSystem.Directory.GetDirectories(targetDirectory);
            foreach (string subdirectory in subdirectoryEntries)
                LoadDirectory(subdirectory);
        }

        protected void LoadFile(string path)
        {
            filePaths.Add(new FileInfo(path));
        }

    }
}
