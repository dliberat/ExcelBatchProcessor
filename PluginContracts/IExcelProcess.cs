using OfficeOpenXml;

namespace PluginContracts
{
    public interface IExcelProcess
    {
        void Run(ExcelWorkbook Workbook);
        void Run(ExcelWorkbook Workbook, object parameters);
    }
}
