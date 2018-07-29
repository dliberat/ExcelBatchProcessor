using PluginContracts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ESPlugins
{
    public class Flatten : IExcelProcess
    {
        public void Run(ExcelWorkbook Workbook)
        {
            if (Workbook.Worksheets.Count < 2) return;

            // Only flatten sheets that have exactly two columns whose headers are "KR" and "JP" (in any order)
            List<KRJPOnlyResult> canFlatten = new List<KRJPOnlyResult>(5);
            foreach (ExcelWorksheet sheet in Workbook.Worksheets)
            {
                KRJPOnlyResult res = IsKRJPOnlyWithHeaders(sheet);
                if (res.CanFlatten) canFlatten.Add(res);
            }
            if (canFlatten.Count < 2) return;

            // Perform the flattening
            ExcelWorksheet flattened = Workbook.Worksheets.Add("Flattened");
            Workbook.Worksheets.MoveToStart(flattened.Name);
            flattened.Cells[1, 1].Value = "KR";
            flattened.Cells[1, 2].Value = "JP";

            foreach (KRJPOnlyResult result in canFlatten)
            {
                ExcelWorksheet sourceSheet = result.Sheet;
                int lastRowSource = sourceSheet.Dimension.End.Row;
                int lastRowTarget = flattened.Dimension.End.Row;

                ExcelRange KRsrc = sourceSheet.Cells[2, result.KRHeaderCol, lastRowSource, result.KRHeaderCol];
                ExcelRange KRtgt = flattened.Cells[lastRowTarget + 1, 1, lastRowTarget + 1 + KRsrc.Rows, 1];

                ExcelRange JPsrc = sourceSheet.Cells[2, result.JPHeaderCol, lastRowSource, result.JPHeaderCol];
                ExcelRange JPtgt = flattened.Cells[lastRowTarget + 1, 2, lastRowTarget + 1 + JPsrc.Rows, 1];

                KRsrc.Copy(KRtgt);
                JPsrc.Copy(JPtgt);

                Workbook.Worksheets.Delete(sourceSheet);
            }
            flattened.Cells.AutoFitColumns();
        }

        private KRJPOnlyResult IsKRJPOnlyWithHeaders(ExcelWorksheet sheet)
        {
            if (sheet.Dimension == null) return new KRJPOnlyResult(sheet, false);
            if (sheet.Dimension.End.Column != 2) return new KRJPOnlyResult(sheet, false);

            int KRHeader = 0;
            int JPHeader = 0;
            try
            {
                KRJPOnlyResult res = new KRJPOnlyResult(sheet);
                string val1 = (string)sheet.Cells[1, 1].Value;
                string val2 = (string)sheet.Cells[1, 2].Value;

                val1 = val1.Trim();
                val2 = val2.Trim();

                if (val1 == "KR") KRHeader = 1;
                if (val1 == "JP") JPHeader = 1;
                if (val2 == "KR") KRHeader = 2;
                if (val2 == "JP") JPHeader = 2;

                res.KRHeaderCol = KRHeader;
                res.JPHeaderCol = JPHeader;
                res.CanFlatten = (KRHeader > 0 && JPHeader > 0);
                return res;
            } catch
            {
                return new KRJPOnlyResult(sheet, false);
            }
        }

        public void Run(ExcelWorkbook Workbook, object parameters)
        {
            Run(Workbook);
        }


    }

    class KRJPOnlyResult
    {
        public ExcelWorksheet Sheet;
        public bool CanFlatten = false;
        public int JPHeaderCol;
        public int KRHeaderCol;

        public KRJPOnlyResult(ExcelWorksheet Sheet, bool CanFlatten = false, int JPCol = 0, int KRCol = 0)
        {
            this.Sheet = Sheet;
            this.CanFlatten = CanFlatten;
            this.JPHeaderCol = JPCol;
            this.KRHeaderCol = KRCol;
        }
    }
}
