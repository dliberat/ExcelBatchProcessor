using ExcelBatchProcessor.Files;
using EBPPluginContracts;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace ExcelBatchProcessor.App
{
    /// <summary>
    /// Processes files, applying transformations and depositing the modified files
    /// in the output directory.
    /// </summary>
    public class ProcessorApp : App
    {
        protected List<IExcelProcess> Plugins;
        protected ConcurrentQueue<FileInfo> FileQueue = new ConcurrentQueue<FileInfo>();
        protected IFileProcessor Processor;

        public ProcessorApp(string Source,
                            string Output,
                            ICollection<IExcelProcess> Plugins,
                            IFileFinder FF,
                            IFileProcessor Processor)
        {
            logger.Debug("Initialized ProcessorApp with source at: {src} and output: {out}", Source, Output);
            this.Source = Source;
            this.Output = Output;
            SourceFiles = new List<FileInfo>(FF.Find());
            this.Plugins = new List<IExcelProcess>(Plugins);
            this.Processor = Processor;
        }
        public override void Run()
        {
            logger.Debug("Running ProcessorApp. Processor count: {p}. Files to process: {f}. Plugins: {plugs}",
                         MaxProcessorCount,
                         SourceFiles.Count,
                         String.Join(", ", Plugins));

            QueueFiles();

            Thread[] workers = new Thread[MaxProcessorCount];
            for (int i = 0; i < MaxProcessorCount; i++)
            {
                workers[i] = new Thread(GetFileFromList);
                workers[i].Start();
            }

            foreach (Thread worker in workers)
                worker.Join();
        }

        private void QueueFiles()
        {
            foreach(FileInfo fi in SourceFiles)
                FileQueue.Enqueue(fi);
        }

        private void GetFileFromList()
        {
            while (FileQueue.TryDequeue(out FileInfo currentFI))
            {
                FileInfo outpath = GetOutputPath(currentFI);
                Processor.Process(currentFI, outpath, Plugins);
            }
                
        }

        /// <summary>
        /// Generates an output path for the modified file. Does not check whether that file name already exists,
        /// so it will overwrite any existing files.
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        private FileInfo GetOutputPath(FileInfo input)
        {
            string outpath = Path.Combine(Output, input.Name);
            return new FileInfo(outpath);
        }
    }
}
