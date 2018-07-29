using PluginContracts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ESPlugins
{
    public class RemoveNonKRJPCols : IExcelProcess
    {
        private const int rowsToSearch = 3;

        public void Run(ExcelWorkbook Workbook)
        {
            foreach (ExcelWorksheet sheet in Workbook.Worksheets)
            {
                // Short-circuit if the sheet is empty or only has 2 columns
                if (sheet.Dimension == null) return;
                if (sheet.Dimension.End.Column < 3) return;

                List<int> toKeep = new List<int>(3);
                bool hasBoth = false;
                for (int row = 1; row <= rowsToSearch; row++)
                {
                    FindResult res = FindKRJPHeaders(sheet, row);
                    if (res.HasBoth)
                    {
                        hasBoth = true;
                        foreach (int i in res.Indices)
                            toKeep.Add(i);
                        break;
                    }
                }

                // Don't delete columns unless we've confirmed that both the KR and JP
                // headers exist
                if (!hasBoth) return;

                for (int column = sheet.Dimension.End.Column; column > 0; column--)
                    if (!toKeep.Contains(column)) sheet.DeleteColumn(column);
            }
        }

        private FindResult FindKRJPHeaders(ExcelWorksheet sheet, int row)
        {
            List<int> indices = new List<int>(3);
            bool hasJP = false;
            bool hasKR = false;

            for (int col = sheet.Dimension.End.Column; col > 0; col--)
            {
                var value = sheet.Cells[row, col].Value;
                if (value is string val)
                {
                    val = val.Trim();

                    if (val == "JP")
                    {
                        hasJP = true;
                        indices.Add(col);
                    }
                    if (val == "KR")
                    {
                        hasKR = true;
                        indices.Add(col);
                    }
                }
            }

            return new FindResult(hasJP && hasKR, indices.ToArray());
        }

        public void Run(ExcelWorkbook Workbook, object parameters)
        {
            Run(Workbook);
        }
    }

    class FindResult
    {
        public bool HasBoth;
        public int[] Indices;

        public FindResult(bool both, int[] indices)
        {
            this.HasBoth = both;
            this.Indices = indices;
        }
    }
}
