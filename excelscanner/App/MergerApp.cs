//using OfficeOpenXml;
//using System;
//using System.Collections.Concurrent;
//using System.Collections.Generic;
//using System.IO;
//using System.Linq;
//using System.Text;
//using System.Threading;
//using System.Threading.Tasks;

//namespace ExcelBatchProcessor
//{
//    public class MergerApp
//    {
//        private string OutputPath;
//        private ConcurrentQueue<FileInfo> FileQueue;
//        private Dictionary<int, int> FileNameCounters;
//        private int PROCESSOR_COUNT = Environment.ProcessorCount;
//        public bool IgnoreNullRows = true;

//        public MergerApp(List<FileInfo> Files, string OutputPath)
//        {
//            FileNameCounters = new Dictionary<int, int>();
//            FileQueue = new ConcurrentQueue<FileInfo>();
//            foreach (FileInfo fi in Files)
//                FileQueue.Enqueue(fi);

//            this.OutputPath = OutputPath;
//        }

//        public void Merge()
//        {
//            Thread[] workers = new Thread[PROCESSOR_COUNT];
//            for (int i = 0; i < PROCESSOR_COUNT; i++)
//            {
//                workers[i] = new Thread(GenerateFiles);
//                FileNameCounters.Add(workers[i].ManagedThreadId, 0);
//                workers[i].Start();
//            }

//            foreach (Thread worker in workers)
//                worker.Join();
//        }

//        private void GenerateFiles()
//        {
//            ExcelPackage package = new ExcelPackage();
//            package.Workbook.Worksheets.Add("Sheet1");
//            try
//            {
//                while (FileQueue.TryDequeue(out FileInfo currentFI))
//                {
//                    AddFileToPackage(package, currentFI);
//                    if (IsFull(package))
//                    {
//                        package.SaveAs(GetFileName());
//                        package.Dispose();
//                        package = new ExcelPackage();
//                        package.Workbook.Worksheets.Add("Sheet1");
//                    }
//                }

//                if (package.Workbook.Worksheets[1].Dimension != null)
//                    package.SaveAs(GetFileName());
//            } catch (Exception ex)
//            {
//                Console.WriteLine("An error occurred and some files could not be merged.");
//                Console.WriteLine(ex.Message);
//                package.Dispose();
//            }
//        }

//        private FileInfo GetFileName()
//        {
//            int id = Thread.CurrentThread.ManagedThreadId;
//            int fileNum = FileNameCounters[id];
//            FileNameCounters[id]++;
//            string name = $"Thread {id} - File {fileNum}.xlsx";
//            return new FileInfo(Path.Combine(OutputPath, name));
//        }

//        private bool IsFull(ExcelPackage package)
//        {
//            ExcelWorkbook wbk = package.Workbook;
//            ExcelWorksheet sheet = wbk.Worksheets[1];
//            if (sheet.Dimension == null) return false;
//            return sheet.Dimension.End.Row > 200;
//        }

//        private void AddFileToPackage(ExcelPackage package, FileInfo currentFI)
//        {
//            ExcelWorksheet targetSheet = package.Workbook.Worksheets[1];
//            int targetSheetStart = targetSheet.Dimension == null ? 1 : targetSheet.Dimension.End.Row + 1;
//            List<object[]> sourceRows = new List<object[]>();

//            using (ExcelPackage source = new ExcelPackage(currentFI))
//            {
//                foreach (ExcelWorksheet sheet in source.Workbook.Worksheets)
//                {
//                    // sheet is empty
//                    if (sheet.Dimension == null) continue;

//                    for (int row = 1; row <= sheet.Dimension.End.Row; row++)
//                    {
//                        object[] rowvalues = new object[sheet.Dimension.End.Column];
//                        for (int col = 1; col <= sheet.Dimension.End.Column; col++)
//                            rowvalues[col - 1] = sheet.Cells[row, col].Value;

//                        bool hasContent = rowvalues.Any(x => x != null);
//                        if ((hasContent || !IgnoreNullRows) && IsNotHeader(rowvalues))
//                            sourceRows.Add(rowvalues);
//                    }
//                }
//            }

//            int currentRow = targetSheetStart;
//            foreach (object[] sourceRow in sourceRows)
//            {
//                for (int col = 0; col < sourceRow.Length; col++)
//                {
//                    targetSheet.Cells[currentRow, col + 1].Value = sourceRow[col];
//                }
//                currentRow++;
//            }
//        }

//        private bool IsNotHeader(object[] rowvalues)
//        {
//            if (rowvalues.Length < 2) return true;
//            if (rowvalues[0] is string str)
//                if (str.Trim() != "KR") return true;
//            if (rowvalues[1] is string str1)
//                if (str1.Trim() != "JP") return true;
//            for (int i = 2; i < rowvalues.Length; i++)
//                if (rowvalues[i] != null) return true;

//            return false;
//        }
//    }
//}
