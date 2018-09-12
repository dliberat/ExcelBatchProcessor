//using PluginContracts;
//using System;
//using OfficeOpenXml;

//namespace ESPlugins
//{
//    /// <summary>
//    /// For some reason this module is corrupting certain files. It may have something to do with
//    /// bad references in the drawing data.
//    /// </summary>
//    public class RemoveImages : IExcelProcess
//    {
//        public void Run(ExcelWorkbook Workbook)
//        {
//            foreach (ExcelWorksheet sheet in Workbook.Worksheets)
//            {
//                try
//                {
//                    if (sheet.Drawings.Count > 0)
//                        sheet.Drawings.Clear();
//                } catch
//                {
//                    Console.WriteLine("Could not remove images from {0}", sheet.Name);
//                }
                
//            }
//        }

//        public void Run(ExcelWorkbook Workbook, object parameters)
//        {
//            Run(Workbook);
//        }
//    }
//}
