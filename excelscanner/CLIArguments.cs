using CommandLine;

namespace ExcelBatchProcessor
{
    class CLIArguments
    {
        [Option('d', "debug", Required = false, HelpText = "Show additional logging output.")]
        public bool Debug { get; set; }

        [Option('i', "input", Required = true, HelpText = "Directory of files to be processed.")]
        public string Input { get; set; }

        [Option('o', "output", Required = true, HelpText = "Destination for processed files.")]
        public string Output { get; set; }

        [Option('f', "file", Default = "", HelpText = "Destination for log files if logging to file.")]
        public string File { get; set; }

        //[Option('p', "plugins", Default = "/plugins", HelpText = "Location for plugin files.")]
        //public string Plugins { get; set; }
    }
}
