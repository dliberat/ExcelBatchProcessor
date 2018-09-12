using NLog;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelBatchProcessor.App
{
    public abstract class App
    {
        /// <summary>
        /// Logger for the application.
        /// </summary>
        protected static Logger logger = LogManager.GetCurrentClassLogger();
        /// <summary>
        /// Location of source files to be processed.
        /// </summary>
        public string Source;
        /// <summary>
        /// Target location for processed files.
        /// </summary>
        public string Output;
        /// <summary>
        /// List of files to be processed.
        /// </summary>
        public List<FileInfo> SourceFiles;
        /// <summary>
        /// Maximum number of processors to use at one time. Defaults to
        /// the maximum number available on the system.
        /// </summary>
        public int MaxProcessorCount = Environment.ProcessorCount;

        public abstract void Run();
    }
}
