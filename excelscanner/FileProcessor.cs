using OfficeOpenXml;
using PluginContracts;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace excelscanner
{
    public class FileProcessor
    {
        // Where to place the files after they have been modified
        public string OutputPath { get; set; }

        // List of filters to apply to each file
        private List<IExcelProcess> Plugins = new List<IExcelProcess>();

        // List of files to be processed
        private ConcurrentQueue<FileInfo> FileQueue;

        public FileProcessor(List<FileInfo> Files, string OutputPath)
        {
            FileQueue = new ConcurrentQueue<FileInfo>();

            foreach (FileInfo fi in Files)
                FileQueue.Enqueue(fi);

            this.OutputPath = OutputPath;
        }

        public void AddPlugin(IExcelProcess Plugin)
        {
            Plugins.Add(Plugin);
            Console.WriteLine("Loaded 1 plugin.");
        }
        public void AddPlugins(ICollection<IExcelProcess> Plugins)
        {
            foreach (IExcelProcess plugin in Plugins)
            {
                this.Plugins.Add(plugin);
            }

            string s = Plugins.Count != 1 ? "s" : "";
            Console.WriteLine($"Loaded {Plugins.Count} plugin{s}.");
        }

        public void Process()
        {
            Thread[] workers = new Thread[Environment.ProcessorCount];
            for (int i = 0; i < Environment.ProcessorCount; i++)
            {
                workers[i] = new Thread(GetFileFromList);
                workers[i].Start();
            }

            foreach (Thread worker in workers)
                worker.Join();

            Console.WriteLine("Completed");
        }

        private void GetFileFromList()
        {
            while (FileQueue.TryDequeue(out FileInfo currentFI))
            {
                ProcessFile(currentFI);
            }
        }

        private void ProcessFile(FileInfo input)
        {
            FileInfo output = GetOutputPath(input);

            if (!input.Extension.Contains("xl"))
            {
                Console.WriteLine("File '{0}' is not an Excel file. Skipping.", input.Name);
                return;
            }

            ExcelPackage package = new ExcelPackage(input);

            foreach (IExcelProcess plugin in Plugins)
                plugin.Run(package.Workbook);

            if (!IsEmptyWorkbook(package.Workbook, package.File.Name))
            {
                package.SaveAs(output);
                Console.WriteLine("Saved file '{0}' to output directory.", package.File.Name);
            } else
            {
                package.Dispose();
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
            string outpath = Path.Combine(OutputPath, input.Name);
            return new FileInfo(outpath);
        }

        private bool IsEmptyWorkbook(ExcelWorkbook Workbook, string FileName)
        {
            foreach (ExcelWorksheet sheet in Workbook.Worksheets)
                if (sheet.Dimension != null) return false;

            Console.WriteLine("Workbook '{0}' is empty. It will not be copied to the output directory.", FileName);
            return true;
        }
    }
}
