using NLog;
using NLog.Config;
using NLog.Targets;
using System;

namespace ExcelBatchProcessor.Logging
{
    /// <summary>
    /// Generates NLog configurations based on arguments provided via the command line.
    /// </summary>
    /// <remarks>
    /// Logging to the console is enabled by default, but only Warn level logs and higher
    /// are shown to the user. In order to enable more detailed logging, the <c>--debug</c>
    /// argument can be passed to the application.
    /// </remarks>
    /// <example>
    /// <code>
    /// C:\>ExcelBatchProcessor.exe -i C:\Users\owner\Desktop\ -o C:\Users\owner\Desktop\processed
    /// // End of output
    /// </code>
    /// </example>
    /// <example>
    /// <code>
    /// C:\>ExcelBatchProcessor.exe -i C:\Users\owner\Desktop\ -o C:\Users\owner\Desktop\processed --debug
    /// 2018-08-14 21:38:17.7392|DEBUG|ExcelBatchProcessor.Program|Initialized logging config with debug: true and file: ""
    /// 2018-08-14 21:38:17.7652|DEBUG|ExcelBatchProcessor.App.App|Initialized ProcessorApp with source at: "C:\Users\owner\Desktop\" and output: "C:\Users\owner\Desktop\processed"
    /// 2018-08-14 21:38:17.7652|INFO|ExcelBatchProcessor.Files.FileFinder|Found 25 files recursively in "C:\Users\owner\Desktop\"
    /// </code>
    /// </example>
    public class ConfigGenerator
    {
        private LoggingConfiguration Config;
        private ConsoleTarget logconsole = new ConsoleTarget("logconsole");

        public ConfigGenerator(bool IsDebugMode, string LogToFile)
        {
            Config = new LoggingConfiguration();

            LogLevel min = IsDebugMode ? LogLevel.Trace : LogLevel.Warn;
            LogLevel max = LogLevel.Fatal;

            Config.AddRule(min, max, logconsole);

            // Fails silently if file path is invalid
            if (LogToFile != "" && Uri.IsWellFormedUriString(LogToFile, UriKind.RelativeOrAbsolute))
            {
                FileTarget logfile = new FileTarget("logfile") { FileName = LogToFile };
                Config.AddRule(min, max, logfile);
            }
        }

        public LoggingConfiguration GetConfig()
        {
            return Config;
        }
    }
}
