using PluginContracts;
using System;
using System.Collections.Generic;
using OfficeOpenXml;
using System.Xml;
using System.Text.RegularExpressions;

namespace ESPlugins
{
    public class RemoveNoKRJPRows : IExcelProcess
    {
        public void Run(ExcelWorkbook Workbook)
        {
            List<ExcelWorksheet> okToDelete = new List<ExcelWorksheet>();

            foreach(ExcelWorksheet sheet in Workbook.Worksheets)
            {
                // This will happen if there is no content in the sheet
                if (sheet.Dimension == null)
                {
                    okToDelete.Add(sheet);
                    continue;
                }

                List<int> rowsToDelete = new List<int>();
                for (int row = sheet.Dimension.End.Row; row > 0; row--)
                    if (!HasJPAndKR(sheet, row)) rowsToDelete.Add(row);

                foreach (int row in rowsToDelete)
                    sheet.DeleteRow(row);
            }

            foreach (ExcelWorksheet sheet in okToDelete)
                if (Workbook.Worksheets.Count > 1) Workbook.Worksheets.Delete(sheet);
        }

        public void Run(ExcelWorkbook Workbook, object parameters)
        {
            Run(Workbook);
        }

        private bool HasJPAndKR(ExcelWorksheet sheet, int row)
        {
            Regex jp = new Regex(@"[\u3041-\u30FF]");
            Regex kr = new Regex(@"[\u1100-\u11FF\uA960-\uA97F\uAC00-\uD7FF]");

            bool hasJP = false;
            bool hasKR = false;
            try
            {
                for (int col = 1; col <= sheet.Dimension.End.Column; col++)
                {
                    var value = sheet.Cells[row, col].Value;
                    if (value == null || !(value is string)) continue;

                    string val = (string)value;

                    // Keep header rows because they help weed out columns later on
                    if (val == "JP" || val == "KR" || val == "EN") return true;

                    if (jp.IsMatch(val)) hasJP = true;
                    if (kr.IsMatch(val)) hasKR = true;
                    if (hasJP && hasKR) return true;
                }

                return false;
            } catch
            {
                return true;
            }
            
        }

    }
}
