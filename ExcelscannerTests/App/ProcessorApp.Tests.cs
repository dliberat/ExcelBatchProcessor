using ExcelBatchProcessor.App;
using ExcelBatchProcessor.Files;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using EBPPluginContracts;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Abstractions;
using System.Threading;

namespace ExcelBatchProcessorTests.App
{
    class MockFileFinder : IFileFinder
    {
        private IFileSystem FileSystem;
        private int FileCount;

        public MockFileFinder(IFileSystem FileSystem, int FileCount = 0)
        {
            this.FileSystem = FileSystem;
            this.FileCount = FileCount;
        }
        public MockFileFinder(int FileCount = 0) : this(new FileSystem(), FileCount)
        {
        }

        public IEnumerable<FileInfo> Find()
        {
            List<FileInfo> list = new List<FileInfo>();

            for (int i = 0; i < FileCount; i++)
            {
                list.Add(new FileInfo($@"c:\bogus\myfile{i}"));
            }

            return list;
        }
    }

    class MockFileProcessor : IFileProcessor
    {
        public int ProcessCalls;

        public void Process(FileInfo FI, FileInfo OutputPath, ICollection<IExcelProcess> Plugins)
        {
            ProcessCalls++;
        }
    }

    class SlowFileProcessor : IFileProcessor
    {
        public void Process(FileInfo FI, FileInfo OutputPath, ICollection<IExcelProcess> Plugins)
        {
            Thread.Sleep(100);
        }
    }

    [TestClass]
    public class ProcessorAppTests
    {
        [TestCategory("App"), TestMethod()]
        public void Constructor_test()
        {
            string source = @"c:\bogus";
            string target = @"d:\fake";
            List<IExcelProcess> plugins = new List<IExcelProcess>();
            IFileFinder ff = new MockFileFinder();
            IFileProcessor fp = new MockFileProcessor();
            ProcessorApp app = new ProcessorApp(source, target, plugins, ff, fp);

            Assert.AreEqual(@"c:\bogus", app.Source);
            Assert.AreEqual(@"d:\fake", app.Output);
            Assert.AreEqual(Environment.ProcessorCount, app.MaxProcessorCount);
            Assert.AreEqual(0, app.SourceFiles.Count);
        }
        [TestCategory("App"), TestMethod()]
        public void Run_No_files_to_process()
        {
            string source = @"c:\bogus";
            string target = @"d:\fake";
            List<IExcelProcess> plugins = new List<IExcelProcess>();
            IFileFinder ff = new MockFileFinder();
            MockFileProcessor fp = new MockFileProcessor();
            ProcessorApp app = new ProcessorApp(source, target, plugins, ff, fp);

            app.Run();
            Assert.AreEqual(0, fp.ProcessCalls);
        }
        [TestCategory("App"), TestMethod()]
        public void Run_process_files()
        {
            string source = @"c:\bogus";
            string target = @"d:\fake";
            List<IExcelProcess> plugins = new List<IExcelProcess>();
            IFileFinder ff = new MockFileFinder(3);
            MockFileProcessor fp = new MockFileProcessor();
            ProcessorApp app = new ProcessorApp(source, target, plugins, ff, fp);

            app.Run();
            Assert.AreEqual(3, fp.ProcessCalls);
        }
        [TestCategory("App"), TestMethod()]
        public void Run_multithreaded()
        {
            // WARNING! This test might fail if normal code execution runs extremely
            // slowly for whatever reason. At any rate, the total running time
            // of the test should be less than the sum of the time taken
            // by the slow file processors

            if (Environment.ProcessorCount < 2)
                return;

            string source = @"c:\bogus";
            string target = @"d:\fake";
            List<IExcelProcess> plugins = new List<IExcelProcess>();
            IFileFinder ff = new MockFileFinder(2);
            SlowFileProcessor fp = new SlowFileProcessor();
            ProcessorApp app = new ProcessorApp(source, target, plugins, ff, fp);

            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();

            app.Run();

            stopWatch.Stop();
            // Get the elapsed time as a TimeSpan value.
            TimeSpan ts = stopWatch.Elapsed;
            Assert.IsTrue(ts.Milliseconds < 150);
        }
    }
}
