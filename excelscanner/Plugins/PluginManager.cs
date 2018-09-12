using EBPPluginContracts;
using System.Collections.Generic;

namespace ExcelBatchProcessor.Plugins
{
    public class PluginManager
    {
        DLLLoader loader;

        public PluginManager(DLLLoader FileLoader)
        {
            loader = FileLoader;
        }

        public ICollection<IExcelProcess> LoadFromDirectory(string path)
        {
            return loader.Load(path);
        }
    }
}
