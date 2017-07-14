using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QA_ReportMergerConsole
{
    internal static partial class Program
    {
        //Pham Thi Ha Linh request
        //theres something wrong...
        private static void CreateTakeRandomCallReport(DataSet col, DataSet outbound, int take, double callLen)
        {
            string folder = string.Format("{0}/{1}", GetCurrentDirectory(), OutputFolder);
            string fileName = string.Format("{0}/{1}", folder, ReportRandomContractName);
            var file = new FileInfo(fileName);
            Directory.CreateDirectory(folder);
            if (File.Exists(fileName))
                File.Delete(fileName);

            using (var package = new ExcelPackage(file))
            {
                var processedCollectorID = new HashSet<string>();
                var processedOutboundIndex = new HashSet<int>();
                var worksheet = package.Workbook.Worksheets.Add("random_contract_report");
                worksheet.OutLineSummaryBelow = false;
                var colTable = col.Tables[0];
                //var outboundTable = RandomOrderTable(outbound.Tables[0]); //random outbound calls
                var outboundTable = outbound.Tables[0];
                //map outbound data
                var outboundDataMap = MapData(outboundTable, OutboundCustomerRefColumn);
                //map col data
                var colMap = MapData(colTable, ColCollectorColumn);
                //set headers
                EPPlusHelper.SetRow(2, 1, GetHeaders(colTable), worksheet, COL_COLOR).AutoFilter = true;
                EPPlusHelper.SetRow(1, 2, GetHeaders(outboundTable), worksheet, OUTBOUND_COLOR);
                int currentRow = 3;
                //loop all col data rows
                for (int rowIndex = 0; rowIndex < colTable.Rows.Count; rowIndex++)
                {
                    var row = colTable.Rows[rowIndex];
                    var collectorCode = row[ColCollectorColumn].ToString();
                    if (processedCollectorID.Contains(collectorCode))
                    {
                        UpdatePerfStat(rowIndex + 1, colTable.Rows.Count);
                        continue;
                    }
                    processedCollectorID.Add(collectorCode);
                    int took = 0;
                    //loop all row related to a collector code
                    foreach (var colMapIndex in colMap[collectorCode])
                    {
                        //contract number of a collector
                        var contractNumber = row[ColContractNoColumn].ToString();
                        //check if this contract number been call (has outbound data)
                        if (outboundDataMap.ContainsKey(string.Format("SHD:{0}", contractNumber)))
                        {
                            //loop all outbound to the related contract
                            foreach (var outboundIndex in outboundDataMap[string.Format("SHD:{0}", contractNumber)])
                            {
                                //skip contract that been processed
                                if (processedOutboundIndex.Contains(outboundIndex)) continue;
                                processedOutboundIndex.Add(outboundIndex);
                                DateTime startDate, endDate;
                                if (DateTime.TryParse(outboundTable.Rows[outboundIndex][OutboundStartDateColumn].ToString(), out startDate) &&
                                    DateTime.TryParse(outboundTable.Rows[outboundIndex][OutboundEndDateColumn].ToString(), out endDate))
                                {
                                    if (endDate.Subtract(startDate) >= TimeSpan.FromSeconds(callLen)) //only take outbound > 30s
                                    {
                                        //write to package
                                        EPPlusHelper.SetRow(currentRow, 1, colTable.Rows[colMapIndex].ItemArray, worksheet, COL_COLOR);
                                        currentRow++;
                                        EPPlusHelper.SetRow(currentRow, 2, outboundTable.Rows[outboundIndex].ItemArray, worksheet, OUTBOUND_COLOR);
                                        currentRow++;
                                        EPPlusHelper.CollapseRow(currentRow - 1, currentRow, worksheet);
                                        took++;
                                        //take only one outbound call for earch contract
                                        break;
                                    }
                                }
                                else
                                    continue;
                            }
                        }
                        else
                            continue;
                        //enough for this collector -> break
                        if (took >= take) break;
                    }
                    //this col data row is done 
                    UpdatePerfStat(rowIndex + 1, colTable.Rows.Count);
                }
                //all done -> save
                WriteLine(string.Empty);
                WriteLine("Writing file...");
                package.Save();
            }
        }

        private static void CreateFullReport(DataSet col, DataSet outbound)
        {
            // Set the file name and get the output directory
            var processedSet = new HashSet<string>();
            string folder = string.Format("{0}/{1}", GetCurrentDirectory(), OutputFolder);
            string fileName = string.Format("{0}/{1}", folder, ReportName);
            var file = new FileInfo(fileName);
            Directory.CreateDirectory(folder);
            if (File.Exists(fileName))
                //delete to override
                File.Delete(fileName);
            int outboundCallCount = 0;
            int contractHasBeenCalledCount = 0;
            using (var package = new ExcelPackage(file))
            {
                // add a new worksheet to the empty workbook
                var worksheet = package.Workbook.Worksheets.Add("merged_report");
                worksheet.OutLineSummaryBelow = false;
                //compose report start
                var colTable = col.Tables[0];
                var outboundTable = outbound.Tables[0];
                //map outbound data
                var outboundDataMap = MapData(outboundTable, OutboundCustomerRefColumn);
                //map col data
                var colMap = MapData(colTable, ColContractNoColumn);
                //set headers
                EPPlusHelper.SetRow(2, 1, GetHeaders(colTable), worksheet, COL_COLOR).AutoFilter = true;
                EPPlusHelper.SetRow(1, 2, GetHeaders(outboundTable), worksheet, OUTBOUND_COLOR);
                int currentRow = 3;
                //int perfIndex = 1;
                for (int i = 0; i < colTable.Rows.Count; i++)
                {
                    string contractNumber = colTable.Rows[i][ColContractNoColumn].ToString();
                    if (string.IsNullOrEmpty(contractNumber) || processedSet.Contains(contractNumber))
                    {
                        UpdatePerfStat(i + 1, colTable.Rows.Count);
                        continue;
                    }
                    //add to skip list
                    processedSet.Add(contractNumber);
                    //find related data in outbound
                    List<int> dataIndexes;
                    if (outboundDataMap.ContainsKey(string.Format("SHD:{0}", contractNumber)))
                    {
                        contractHasBeenCalledCount++;
                        dataIndexes = outboundDataMap[string.Format("SHD:{0}", contractNumber)];
                    }
                    else
                    {
                        //no related data in outbound -> add later at the end
                        _addLaterRows.Add(colTable.Rows[i]);
                        UpdatePerfStat(i + 1, colTable.Rows.Count);
                        continue;
                    }
                    ////write col data to sheet
                    int rowToCollapseFrom = 0;
                    foreach (int contractIndex in colMap[contractNumber])
                    {
                        if (contractIndex == i)
                        //first one of each contract has no padding
                        {
                            rowToCollapseFrom = currentRow + 1;
                            EPPlusHelper.SetRow(currentRow, 1, colTable.Rows[contractIndex].ItemArray, worksheet,
                                COL_COLOR);
                        }
                        else
                        {
                            EPPlusHelper.SetRow(currentRow, 2, colTable.Rows[contractIndex].ItemArray, worksheet,
                                OUTBOUND_COLOR);
                        }
                        currentRow++;
                    }
                    //write outbound data to sheet
                    foreach (int index in dataIndexes)
                    {
                        outboundCallCount++;
                        EPPlusHelper.SetRow(currentRow, 2, outboundTable.Rows[index].ItemArray, worksheet,
                            OUTBOUND_COLOR);
                        currentRow++;
                    }
                    if (rowToCollapseFrom == 0)
                        throw new InvalidOperationException("rowToCollapseFrom has not been set!");
                    EPPlusHelper.CollapseRow(rowToCollapseFrom, currentRow, worksheet);
                    UpdatePerfStat(i + 1, colTable.Rows.Count);
                }
                //add back empty ones
                foreach (var noRelatedDataRow in _addLaterRows)
                {
                    EPPlusHelper.SetRow(currentRow, 1, noRelatedDataRow.ItemArray, worksheet, COL_COLOR);
                    currentRow++;
                }
                //compose report end
                WriteLine(string.Empty);
                WriteLine("Writing file...");
                WriteLine(string.Empty);
                WriteLine("Summary:");
                WriteLine(string.Format("Contract has been called: {0}", contractHasBeenCalledCount));
                WriteLine(string.Format("Total outbound calls: {0}", outboundCallCount));
                package.Save();
            }
        }
    }
}
