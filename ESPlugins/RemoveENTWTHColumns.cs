using PluginContracts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Text.RegularExpressions;

namespace ESPlugins
{
    public class RemoveENTWTHColumns : IExcelProcess
    {
        // How many rows to look in to try to identify the headers
        private int rowsToCheck = 3;

        private string[] deletionTargets = new string[] { "EN", "TW", "TH" };

        public void Run(ExcelWorkbook Workbook)
        {
            foreach (ExcelWorksheet sheet in Workbook.Worksheets)
            {
                if (sheet.Dimension == null) continue;

                List<int> targetCols = new List<int>();
                // Iterate backwards so that column numbers go in reverse order
                for (int col = sheet.Dimension.End.Column; col > 0; col--)
                    if (ContainsTargetText(sheet, col)) targetCols.Add(col);

                // Need to delete in reverse order because column numbers change with each deletion
                foreach (int column in targetCols)
                    sheet.DeleteColumn(column);
            }
        }

        public void Run(ExcelWorkbook Workbook, object parameters)
        {
            Run(Workbook);
        }

        private bool ContainsTargetText(ExcelWorksheet sheet, int col)
        {
            for (int row = 1; row <= rowsToCheck; row++)
            {
                try
                {
                    string val = (string)sheet.Cells[row, col].Value;
                    if (deletionTargets.Contains(val)) return true;
                } catch
                {
                    continue;
                }
            }

            return false;
        }
    }
}
