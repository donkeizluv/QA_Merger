//using OfficeOpenXml;
//using System;
//using System.Collections.Concurrent;
//using System.Collections.Generic;
//using System.Data;
//using System.IO;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

//namespace QA_ReportMergerConsole
//{
//    partial class Program
//    {
//        //does not seem to be faster at all.
//        //skipping same contracts
//        private static async void CreateExcelReport(DataSet col, DataSet outbound, int threads)
//        {
//            var fileName = string.Format("{0}/{1}", GetCurrentDirectory(), "merged_report.xlsx");
//            var file = new FileInfo(fileName);
//            if (File.Exists(fileName))
//                File.Delete(fileName);
//            using (var package = new ExcelPackage(file))
//            {
//                var worksheet = package.Workbook.Worksheets.Add("merged_report");
//                worksheet.OutLineSummaryBelow = false;
//                var colTable = col.Tables[0];
//                var outboundTable = outbound.Tables[0];

//                //set headers
//                EPPlusHelper.SetRow(2, 1, GetHeaders(colTable), worksheet, COL_COLOR).AutoFilter = true;
//                EPPlusHelper.SetRow(1, 2, GetHeaders(outboundTable), worksheet, OUTBOUND_COLOR);
//                //mapping
//                var outboundDataMap = MapData(outboundTable, outbound_report_column);
//                //create work
//                var taskList = new List<Task>();
//                for (int i = 0; i < threads; i++)
//                {
//                    int from = i * (colTable.Rows.Count / threads);
//                    int to;
//                    if (i < threads - 1)
//                        to = (i + 1) * (colTable.Rows.Count / threads);
//                    else //last thread
//                        to = colTable.Rows.Count - 1; //solve int conversion problem

//                    WriteLine(string.Format("start thread {0}: {1} - {2}", i + 1, from + 1, to + 1)); //convert from index for easy understanding
//                    var task = new Task(() => { Work(from, to, colTable, outboundDataMap, outboundTable, worksheet); });
//                    taskList.Add(task);
//                    task.Start();

//                }
//                //wait all tasks to complete
//                WriteLine("Waiting for threads to complete.");
//                Task.WaitAll(taskList.ToArray());
//                WriteLine("all tasks completed!");
//                //add back empty ones
//                foreach (var noRelatedDataRow in _addLaterBag)
//                {
//                    EPPlusHelper.SetRow(_currentRow, 1, noRelatedDataRow.ItemArray, worksheet, COL_COLOR);
//                    _currentRow++;
//                }
//                WriteLine("Writing file...");
//                package.Save();
//            }
//        }
//        static ConcurrentBag<string> _excludeBag = new ConcurrentBag<string>();
//        static ConcurrentBag<DataRow> _addLaterBag = new ConcurrentBag<DataRow>();
//        static int _currentRow = 3; //multi threaded index reminder
//        static object _excelWriterLock = new object();
//        static int _perfCounter = 1;
//        //static object _perfCounterLock = new object();
//        static void Work(int from, int to, DataTable colData, Dictionary<string, List<int>> outboundMap, DataTable outboundData, ExcelWorksheet ws)
//        {
//            //RegisterThread(from, to);
//            //WriteLine("Thread started...");
//            for (int i = from; i < to; i++)
//            {
//                var row = colData.Rows[i];
//                string no = row[col_report_column].ToString();
//                if (string.IsNullOrEmpty(no) || _excludeBag.Contains(no))
//                {
//                    continue;
//                }
//                _excludeBag.Add(no);
//                //find related data in outbound
//                List<int> dataIndexes;
//                if (outboundMap.ContainsKey(string.Format("SHD:{0}", no)))
//                    dataIndexes = outboundMap[string.Format("SHD:{0}", no)];
//                else
//                {
//                    //no related data in outbound -> add later at the end
//                    _addLaterBag.Add(row);
//                    //UpdatePerfStat(perfIndex, colTable.Rows.Count);
//                    //perfIndex++;
//                    continue;
//                }
//                //maybe better to use increment lock
//                lock (_excelWriterLock)
//                {
//                    //write col data to sheet
//                    EPPlusHelper.SetRow(_currentRow, 1, row.ItemArray, ws, COL_COLOR);
//                    //advance next line
//                    _currentRow++;
//                    //write outbound data to sheet
//                    int rowToCollapseFrom = _currentRow;
//                    foreach (var index in dataIndexes)
//                    {
//                        EPPlusHelper.SetRow(_currentRow, 2, outboundData.Rows[index].ItemArray, ws, OUTBOUND_COLOR);
//                        _currentRow++;
//                    }
//                    EPPlusHelper.CollapseRow(rowToCollapseFrom, _currentRow - 1, ws);
//                    //UpdatePerfStat(perfIndex, colTable.Rows.Count);
//                    //perfIndex++;
//                }
//                //Interlocked.Increment(ref _perfCounter);
//                //UpdatePerfStat(_perfCounter, colData.Rows.Count);
//            }
//        }
//    }
//}

