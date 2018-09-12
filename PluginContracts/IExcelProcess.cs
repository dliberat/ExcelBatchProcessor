using OfficeOpenXml;

namespace EBPPluginContracts
{
    public interface IExcelProcess
    {
        void Run(string Path, ExcelWorkbook Workbook);
    }
}
