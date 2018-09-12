//using PluginContracts;
//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using OfficeOpenXml;

//namespace ESPlugins
//{
//    public class RemoveEmptySheets : IExcelProcess
//    {
//        public void Run(ExcelWorkbook Workbook)
//        {
//            List<ExcelWorksheet> toDelete = new List<ExcelWorksheet>();
//            foreach (ExcelWorksheet sheet in Workbook.Worksheets)
//                if (sheet.Dimension == null) toDelete.Add(sheet);

//            foreach (ExcelWorksheet sheet in toDelete)
//                if (Workbook.Worksheets.Count > 1) Workbook.Worksheets.Delete(sheet);
//        }

//        public void Run(ExcelWorkbook Workbook, object parameters)
//        {
//            Run(Workbook);
//        }
//    }
//}
