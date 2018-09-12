using OfficeOpenXml;
using EBPPluginContracts;
using System.Collections.Generic;
using System.IO;
using System.IO.Abstractions;
using NLog;
using System;

namespace ExcelBatchProcessor.App
{
    public class BasicFileProcessor : IFileProcessor
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();

        readonly IFileSystem FileSystem;

        /// <summary>
        /// Constructor for the <see cref="BasicFileProcessor"/> class. This constructor is mainly
        /// reserved for testing purposes. Consumers of the class should use the
        /// <see cref="BasicFileProcessor()"/> constructor instead.
        /// </summary>
        /// <param name="FileSystem"></param>
        public BasicFileProcessor(IFileSystem FileSystem)
        {
            this.FileSystem = FileSystem;
        }
        /// <summary>
        /// Default constructor for the <see cref="BasicFileProcessor"/> class.
        /// </summary>
        public BasicFileProcessor() : this(FileSystem: new FileSystem())
        {

        }

        public void Process(FileInfo InputPath, FileInfo OutputPath, ICollection<IExcelProcess> Plugins)
        {
            if (!InputPath.Extension.Contains("xl"))
            {
                logger.Warn("File '{0}' is not an Excel file. Skipping.", InputPath.Name);
                return;
            }

            ExcelPackage package = new ExcelPackage(InputPath);

            foreach (IExcelProcess plugin in Plugins)
            {
                try
                {
                    plugin.Run(InputPath.FullName, package.Workbook);
                } catch (Exception ex)
                {
                    logger.Error(ex, "An exception occurred while applying the '{type}' " + 
                        "plugin to the file '{f}'.", plugin.GetType(), InputPath.FullName);
                }
            }
                
            // TODO: Add a pre-save hook (eg., to avoid saving empty files, etc.)
            package.SaveAs(OutputPath);
            logger.Debug("Saved file '{0}' to output directory.", package.File.Name);
        }

        //private bool IsEmptyWorkbook(ExcelWorkbook Workbook, string FileName)
        //{
        //    foreach (ExcelWorksheet sheet in Workbook.Worksheets)
        //        if (sheet.Dimension != null) return false;

        //    Console.WriteLine("Workbook '{0}' is empty. It will not be copied to the output directory.", FileName);
        //    return true;
        //}
    }
}
