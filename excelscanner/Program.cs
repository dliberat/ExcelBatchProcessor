using PluginContracts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelscanner
{
    class Program
    {
        static string workingDir = @"C:\Users\owner\Desktop\e"; //Environment.CurrentDirectory
        static string outputDir = @"C:\Users\owner\Desktop\e_output";

        static void Main(string[] args)
        {
            string[] path = new string[1] { GetDirPath() };
            FileFinder ff = new FileFinder(path);
            List<FileInfo> files = ff.LoadFiles();

            Console.WriteLine($"Loaded {files.Count} files. Press ENTER to process, or type MERGE to merge.");
            string ret = Console.ReadLine();

            while (true)
            {
                if (ret.ToLower() == "merge")
                {
                    Merge(files);
                    break;
                } else if (ret == "") {
                    Process(files);
                    break;
                } else
                {
                    Console.WriteLine("Invalid input.");
                    ret = Console.ReadLine();
                }
            }

            Console.ReadLine();
        }

        static void Merge(List<FileInfo> files)
        {
            Merger am = new Merger(files, outputDir);
            am.Merge();

            Console.WriteLine("Merger complete.");
        }

        static void Process(List<FileInfo> files)
        {
            ICollection<IExcelProcess> plugins = LoadPlugins();

            FileProcessor fp = new FileProcessor(files, outputDir);
            fp.AddPlugins(plugins);

            Console.WriteLine("File Processor ready. Press ENTER to begin processing.");
            Console.ReadLine();

            fp.Process();

            Console.WriteLine("Processing finished.");
        }

        static string GetDirPath()
        {
            string curpath = workingDir;
            Console.WriteLine($"Currently working in: {curpath}\nENTER to use current path, or type in a path.");

            while (true)
            {
                string retval = Console.ReadLine();
                if (retval == "")
                {
                    return curpath;
                }
                else if (Directory.Exists(retval))
                {
                    return retval;
                }
                else
                {
                    Console.WriteLine("Not a valid directory");
                }
            }
        }

        private static ICollection<IExcelProcess> LoadPlugins()
        {
            string cwd = Directory.GetCurrentDirectory();
            string pluginDir = Path.Combine(cwd, "plugins");
            PluginLoader loader = new PluginLoader();
            ICollection<IExcelProcess> plugins = loader.Load(pluginDir);

            // Drawing removal not working reliably. May be related to
            // drawings with identical names or corrupted references
            // plugins.Add(new ESPlugins.RemoveImages());

            // Remove all hidden rows, empty rows, empty columns, and hidden sheets
            plugins.Add(new ESPlugins.RemoveEmptyHidden());

            // Remove rows that don't include BOTH KR and JP text
            // Removes rows with only numerical data too, since presumably these don't
            // really need to be put into TM
            plugins.Add(new ESPlugins.RemoveNoKRJPRows());

            //plugins.Add(new ESPlugins.RemoveENTWTHColumns());
            plugins.Add(new ESPlugins.RemoveNonKRJPCols());

            plugins.Add(new ESPlugins.Flatten());

            plugins.Add(new ESPlugins.RemoveEmptySheets());
            return plugins;
        }
    }
}
