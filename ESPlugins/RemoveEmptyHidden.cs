using OfficeOpenXml;
using PluginContracts;
using System.Collections.Generic;
using System;

namespace ESPlugins
{
    public class RemoveEmptyHidden : IExcelProcess
    {
        public void Run(ExcelWorkbook Workbook)
        {
            List<ExcelWorksheet> okToDelete = new List<ExcelWorksheet>();

            foreach (ExcelWorksheet sheet in Workbook.Worksheets)
            {
                // Dimension==null will happen if there is no content in the sheet
                if (sheet.Dimension == null || sheet.Hidden != eWorkSheetHidden.Visible)
                {
                    okToDelete.Add(sheet);
                    continue;
                }

                // Rows
                List<int> rowsToDelete = new List<int>();
                for (int row = sheet.Dimension.End.Row; row > 0; row --)
                    if (IsEmptyOrHiddenRow(sheet, row)) rowsToDelete.Add(row);
                foreach (int row in rowsToDelete)
                    sheet.DeleteRow(row);

                if (sheet.Dimension == null)
                {
                    okToDelete.Add(sheet);
                    continue;
                }

                // Cols
                List<int> colsToDelete = new List<int>();
                for (int col = sheet.Dimension.End.Column; col > 0; col--)
                    if (IsEmptyCol(sheet, col)) colsToDelete.Add(col);
                foreach (int col in colsToDelete)
                    sheet.DeleteColumn(col);
            }

            foreach (ExcelWorksheet sheet in okToDelete)
            {
                if (Workbook.Worksheets.Count > 1)
                    Workbook.Worksheets.Delete(sheet);
            }
        }

        public void Run(ExcelWorkbook Workbook, object parameters)
        {
            Run(Workbook);
        }

        private bool IsEmptyOrHiddenRow(ExcelWorksheet sheet, int row)
        {
            ExcelRow theRow = sheet.Row(row);
            if (theRow.Hidden)
                return true;

            for (int col = 1; col <= sheet.Dimension.End.Column; col++)
            {
                try
                {
                    var val = sheet.Cells[row, col].Value;
                    if (val != null) return false;
                } catch
                {
                    // safety catch. Assume there's content in the row
                    return false;
                }
            }
            return true;
        }

        private bool IsEmptyCol(ExcelWorksheet sheet, int col)
        {
            for (int row = 1; row < sheet.Dimension.End.Row; row++)
            {
                try
                {
                    var val = sheet.Cells[row, col].Value;
                    if (val != null) return false;
                }
                catch
                {
                    return false;
                }
            }
            return true;
        }
    }
}
