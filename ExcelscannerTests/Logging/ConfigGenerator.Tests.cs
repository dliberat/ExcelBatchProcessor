using ExcelBatchProcessor.Logging;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NLog.Config;
using System.Linq;

namespace ExcelBatchProcessorTests.Logging
{
    [TestClass]
    public class ConfigGeneratorTests
    {
        [TestCategory("Logging"), TestMethod()]
        public void TestConstructor_ConsoleOnly()
        {
            ConfigGenerator generator = new ConfigGenerator(false, "");
            LoggingConfiguration config = generator.GetConfig();
            string[] targetNames = config.AllTargets.Select(x => x.Name).ToArray();

            Assert.AreEqual("logconsole", targetNames[0]);
            Assert.AreEqual(1, targetNames.Length);
        }
        [TestCategory("Logging"), TestMethod()]
        public void TestConstructor_ConsoleAndFile()
        {
            ConfigGenerator generator = new ConfigGenerator(false, "myFile.log");
            LoggingConfiguration config = generator.GetConfig();
            string[] targetNames = config.AllTargets.Select(x => x.Name).ToArray();

            Assert.IsTrue(targetNames.Contains("logconsole"));
            Assert.IsTrue(targetNames.Contains("logfile"));
        }
        [TestCategory("Logging"), TestMethod()]
        public void TestConstructor_BadFile()
        {
            ConfigGenerator generator = new ConfigGenerator(false, "C:[/\broken");
            LoggingConfiguration config = generator.GetConfig();
            string[] targetNames = config.AllTargets.Select(x => x.Name).ToArray();

            Assert.AreEqual("logconsole", targetNames[0]);
            Assert.AreEqual(1, targetNames.Length);
        }
        [TestCategory("Logging"), TestMethod()]
        public void TestConstructor_DebugMode()
        {
            ConfigGenerator prod_generator = new ConfigGenerator(false, "");
            LoggingConfiguration config = prod_generator.GetConfig();
            LoggingRule rule = config.LoggingRules[0];
            string[] levels = rule.Levels.Select(x=>x.Name).ToArray();

            Assert.AreEqual(3, levels.Length);
            Assert.IsTrue(levels.Contains("Warn"));
            Assert.IsTrue(levels.Contains("Error"));
            Assert.IsTrue(levels.Contains("Fatal"));

            ConfigGenerator debug_generator = new ConfigGenerator(true, "");
            LoggingConfiguration debug_config = debug_generator.GetConfig();
            LoggingRule debug_rule = debug_config.LoggingRules[0];
            string[] debug_levels = debug_rule.Levels.Select(x => x.Name).ToArray();

            Assert.AreEqual(6, debug_levels.Length);
            Assert.IsTrue(debug_levels.Contains("Trace"));
            Assert.IsTrue(debug_levels.Contains("Debug"));
            Assert.IsTrue(debug_levels.Contains("Info"));
            Assert.IsTrue(debug_levels.Contains("Warn"));
            Assert.IsTrue(debug_levels.Contains("Error"));
            Assert.IsTrue(debug_levels.Contains("Fatal"));
        }
    }
}
