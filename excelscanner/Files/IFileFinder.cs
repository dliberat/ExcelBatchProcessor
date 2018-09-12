using System.Collections.Generic;
using System.IO;

namespace ExcelBatchProcessor.Files
{
    public interface IFileFinder
    {
        IEnumerable<FileInfo> Find();
    }
}
