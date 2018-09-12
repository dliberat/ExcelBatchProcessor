using EBPPluginContracts;
using OfficeOpenXml;

namespace ESPlugins
{
    public class DoNothingPlugin : IExcelProcess
    {
        public void Run(string Path, ExcelWorkbook Workbook)
        {
        }
    }
}
