using EBPPluginContracts;
using System;
using System.Collections.Generic;
using System.IO;
using CommandLine;
using NLog;
using System.Diagnostics.CodeAnalysis;

namespace ExcelBatchProcessor
{
    [ExcludeFromCodeCoverage]
    class Program
    {
        static Logger logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// Entry point for the application.  
        /// </summary>
        /// <param name="args">Command-line arguments. Recognized parameters include <c>--debug</c>,
        /// <c>--input</c>, and <c>--output</c>.</param>
        public static void Main(string[] args)
        {
            ParserResult<CLIArguments> pargs = Parser.Default.ParseArguments<CLIArguments>(args);

            pargs.WithNotParsed<CLIArguments>(a => Environment.Exit(0));

            pargs.WithParsed<CLIArguments>(a => BootstrapApp(a));
        }
        private static void BootstrapApp(CLIArguments pargs)
        {
            // log config
            Logging.ConfigGenerator logConfig = new Logging.ConfigGenerator(pargs.Debug, pargs.File);
            LogManager.Configuration = logConfig.GetConfig();
            logger.Debug("Initialized logging config with debug: {d} and file: {f}", pargs.Debug, pargs.File);

            // plugins
            string cwd = Directory.GetCurrentDirectory();
            string pluginDir = Path.Combine(cwd, "plugins");
            Plugins.DLLLoader loader = new Plugins.DLLLoader();
            // Inject the loader, then use the manager to add in additional/default
            // plugins as necessary
            Plugins.PluginManager manager = new Plugins.PluginManager(loader);
            ICollection<IExcelProcess> plugins = manager.LoadFromDirectory(pluginDir);

            // main app
            Files.FileFinder ff = new Files.FileFinder(new string[] { pargs.Input });
            App.IFileProcessor fp = new App.BasicFileProcessor();
            App.App mainApp = new App.ProcessorApp(pargs.Input, pargs.Output, plugins, ff, fp);
            mainApp.Run();
        }
    }
}
