using ExcelBatchProcessor.Plugins;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using EBPPluginContracts;
using System.Collections.Generic;
using System.IO;
using System.IO.Abstractions.TestingHelpers;

namespace ExcelBatchProcessorTests.Plugins
{
    [TestClass]
    public class DLLLoaderTests
    {
        [TestCategory("Plugins"), TestMethod()]
        public void Test_Nothing_To_Load()
        {
            var fileSystem = new MockFileSystem(new Dictionary<string, MockFileData>
            {
                { @"c:\myfile.txt", new MockFileData("bogus!") },
            });

            DLLLoader loader = new DLLLoader(fileSystem);
            ICollection<IExcelProcess> plugins = loader.Load(@"C:\");
            Assert.AreEqual(0, plugins.Count);
        }
        [TestCategory("Plugins"), TestMethod()]
        public void Test_Null_Argument()
        {
            var fileSystem = new MockFileSystem(new Dictionary<string, MockFileData>(0));

            DLLLoader loader = new DLLLoader();
            ICollection<IExcelProcess> plugins = loader.Load(null);
            Assert.AreEqual(0, plugins.Count);
        }
        [TestCategory("Plugins"), TestMethod()]
        public void Test_Load_Assembly()
        {
            // Plugin stub
            string cwd = Directory.GetCurrentDirectory();
            string stubs_directory = Path.Combine(cwd, "../stubs/");

            DLLLoader loader = new DLLLoader();
            ICollection<IExcelProcess> plugins = loader.Load(stubs_directory);
            Assert.AreEqual(1, plugins.Count);
        }
    }
}
