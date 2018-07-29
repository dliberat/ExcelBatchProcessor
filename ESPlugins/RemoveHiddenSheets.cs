using System;
using OfficeOpenXml;
using PluginContracts;
using System.Collections.Generic;

namespace ESPlugins
{
    public class RemoveHiddenSheets : IExcelProcess
    {
        public void Run(ExcelWorkbook Workbook)
        {
            List<ExcelWorksheet> hiddenSheets = new List<ExcelWorksheet>();

            foreach (ExcelWorksheet sheet in Workbook.Worksheets)
            {
                if (sheet.Hidden != eWorkSheetHidden.Visible)
                    hiddenSheets.Add(sheet);
            }
            foreach (ExcelWorksheet sheet in hiddenSheets)
                Workbook.Worksheets.Delete(sheet);

            if (hiddenSheets.Count > 0)
                Console.WriteLine("Removed {0} hidden sheets from package.", hiddenSheets.Count);
        }

        public void Run(ExcelWorkbook Workbook, object parameters)
        {
            Run(Workbook);
        }
    }
}
