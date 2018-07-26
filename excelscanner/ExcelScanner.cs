using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace excelscanner
{
    public class ExcelScanner
    {
        private FileInfo fi;
        private Regex rgKR = new Regex(@"KR|KO", RegexOptions.IgnoreCase);
        private Regex rgJP = new Regex(@"JP", RegexOptions.IgnoreCase);

        private int JPHeaderCol = -1;
        private int KRHeaderCol = -1;

        public bool Modified = false;

        public ExcelScanner(FileInfo FI)
        {
            fi = FI;
        }

        public void DeleteNonKRJPHeaders()
        {
            using (ExcelPackage package = new ExcelPackage(fi))
            {
                foreach (ExcelWorksheet sheet in package.Workbook.Worksheets)
                {
                    bool hasHeaders = false;
                    try
                    {
                        hasHeaders = HasHeaders(sheet);
                    }
                    catch
                    {
                        return;
                    }

                    if (hasHeaders)
                    {
                        Modified = true;
                        for (int col = sheet.Dimension.Columns; col > 0; col--)
                        {
                            if (col != JPHeaderCol && col != KRHeaderCol)
                            {
                                sheet.DeleteColumn(col);
                            }
                        }

                        package.SaveAs(fi);
                    }
                }
            }
        }

        private bool HasHeaders(ExcelWorksheet sheet)
        {
            // scan the first row
            int colCount = sheet.Dimension.Columns;
            int row = 1;
            for (int col = 1; col <= colCount; col ++)
            {
                string header = (string)sheet.Cells[row, col].Value;
                bool isJP = rgJP.IsMatch(header);
                if (isJP && JPHeaderCol > -1) return false;
                if (isJP) JPHeaderCol = col;

                bool isKR = rgKR.IsMatch(header);
                if (isKR && KRHeaderCol > -1) return false;
                if (isKR) KRHeaderCol = col;
            }

            if (JPHeaderCol < 0 || KRHeaderCol < 0)
            {
                Console.WriteLine("\t{0} | Could not identify headers.", fi.Name);
                return false;
            }

            Console.WriteLine("\t{0} | Found headers at cols JP: {1}, KR: {2}", fi.Name, JPHeaderCol, KRHeaderCol);
            return false;
        }
    }
}
