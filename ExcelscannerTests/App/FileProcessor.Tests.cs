using Microsoft.VisualStudio.TestTools.UnitTesting;
using EBPPluginContracts;
using System;
using System.Collections.Generic;
using OfficeOpenXml;
using System.IO.Abstractions.TestingHelpers;
using ExcelBatchProcessor.App;
using System.IO;

namespace ExcelBatchProcessorTests.App
{
    class MockExcelProcess : IExcelProcess
    {
        public int RunCalls = 0;
        public void Run(string Path, ExcelWorkbook Workbook)
        {
            RunCalls++;
        }
        public void Run(ExcelWorkbook Workbook, object parameters)
        {
        }
    }

    [TestClass]
    public class BasicFileProcessorTests
    {
        [TestCategory("App"), TestMethod()]
        public void BasicFileProcessor_NonExcel_Files()
        {
            MockExcelProcess plugin = new MockExcelProcess();
            List<IExcelProcess> plugins = new List<IExcelProcess>()
            {
                plugin,
            };

            var fileSystem = new MockFileSystem(new Dictionary<string, MockFileData>
            {
                { @"c:\myfile.txt", new MockFileData("bogus!") },
            });

            BasicFileProcessor processor = new BasicFileProcessor(fileSystem);
            FileInfo inputFile = new FileInfo(@"c:\myfile.txt");
            FileInfo outputFile = new FileInfo(@"c:\myfile-done.txt");
            processor.Process(inputFile, outputFile, plugins);
            Assert.AreEqual(0, plugin.RunCalls);
        }
        [TestCategory("App"), TestMethod()]
        public void BasicFileProcessor_Excel_Files()
        {
            MockExcelProcess plugin = new MockExcelProcess();
            List<IExcelProcess> plugins = new List<IExcelProcess>()
            {
                plugin,
            };

            // Stub excel file
            string cwd = Directory.GetCurrentDirectory();
            string stubs_directory = Path.Combine(cwd, "../stubs/");
            string inputpath = Path.Combine(stubs_directory, "sample.xlsx");
            string outputpath = Path.Combine(stubs_directory, "sample-done.xlsx");

            BasicFileProcessor processor = new BasicFileProcessor();
            FileInfo inputFile = new FileInfo(inputpath);
            FileInfo outputFile = new FileInfo(outputpath);
            processor.Process(inputFile, outputFile, plugins);

            bool file_was_created = File.Exists(outputpath);

            if (file_was_created)
            {
                try
                {
                    File.Delete(outputpath);
                } catch
                {
                    Console.WriteLine("ERROR: Could not delete file '{0}'", outputpath);
                }
            }
            Assert.IsTrue(file_was_created);
            Assert.AreEqual(1, plugin.RunCalls);
        }
    }
}
