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

        static void Main(string[] args)
        {
            string[] path = new string[1] { GetDirPath() };
            FileProcessor fp = new FileProcessor();
            fp.Process(path);

            Console.ReadLine();
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
    }
}
