using ExcelBatchProcessor.Plugins;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using EBPPluginContracts;
using System.Collections.Generic;

namespace ExcelBatchProcessorTests.Plugins
{
    [TestClass]
    public class PluginManagerTests
    {
        class DLLLoaderStub : DLLLoader
        {
            public int loadCalls;
            public override ICollection<IExcelProcess> Load(string directory)
            {
                loadCalls++;
                return new List<IExcelProcess>(0);
            }
        }

        [TestCategory("Plugins"), TestMethod()]
        public void PluginManager_Constructor()
        {

            DLLLoaderStub loader = new DLLLoaderStub();
            PluginManager manager = new PluginManager(loader);
            manager.LoadFromDirectory("foo");
            Assert.AreEqual(1, loader.loadCalls);
        }
    }
}
