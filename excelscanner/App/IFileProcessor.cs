using EBPPluginContracts;
using System.Collections.Generic;
using System.IO;

namespace ExcelBatchProcessor.App
{
    public interface IFileProcessor
    {
        void Process(FileInfo FI, FileInfo OutputPath, ICollection<IExcelProcess> Plugins);
    }
}
