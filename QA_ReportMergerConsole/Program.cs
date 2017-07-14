using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using QA_ReportMergerConsole.Reader;
using QA_ReportMergerConsole;

namespace QA_ReportMergerConsole
{
    internal partial class Program
    {
        //configurable
        private static int randomReportTake = 5;
        private static int conversationLen = 30;
        //readonly
        private static readonly Color COL_COLOR = Color.FromArgb(130, 197, 130);
        private static readonly Color OUTBOUND_COLOR = Color.FromArgb(220, 201, 100);
        //columns
        private const string ColContractNoColumn = "Agreement No";
        private const string ColCollectorColumn = "Collector";
        private const string OutboundCustomerRefColumn = "Customer Ref No";
        private const string OutboundStartDateColumn = "StartDateTime";
        private const string OutboundEndDateColumn = "EndDateTime";

        //file names
        private const string ColFolder = "col_data";
        private const string OutboundFolder = "outbound_data";
        private const string OutputFolder = "output";

        //Pham Thi Ha Linh's reports
        private const string ReportName = "merged_report.xlsx";
        private const string ReportRandomContractName = "random_contract_report.xlsx";

        //Pham Thi Phuong Nga's reports
        //private const string ReportName = "merged_report.xlsx";

        private static readonly List<DataRow> _addLaterRows = new List<DataRow>();

        //private static void LoadAssembly()
        //{
        //    AppDomain.CurrentDomain.AssemblyResolve += (sender, args) => {
        //        String resourceName = new AssemblyName(args.Name).Name + ".dll";
        //        using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
        //        {
        //            Byte[] assemblyData = new Byte[stream.Length];
        //            stream.Read(assemblyData, 0, assemblyData.Length);
        //            return Assembly.Load(assemblyData);
        //        }
        //    };
        //}
        private static void Main(string[] args)
        {
            //LoadAssembly();
            Console.ForegroundColor = ConsoleColor.White;
            //col = xlsx
            //outbound = csv
            WriteLine("--------------------------------------------------------------", ConsoleColor.Yellow);
            WriteLine("----------------------*-*-----------*-*-----------------------", ConsoleColor.Yellow);
            WriteLine("-------------------*-------*-----*-------*--------------------", ConsoleColor.Yellow);
            WriteLine("------------------*-----------*-----------*-------------------", ConsoleColor.Yellow);
            WriteLine("------------------*- QA Report Merger v1.1*-------------------", ConsoleColor.Yellow);
            WriteLine("-------------------*----- HD SAISON -----*--------------------", ConsoleColor.Yellow);
            WriteLine("--------------------*---IT Department---*--------------------", ConsoleColor.Yellow);
            WriteLine("---------------------*-------000-------*-----------------------", ConsoleColor.Yellow);
            WriteLine("------------------------*-----------*-------------------------", ConsoleColor.Yellow);
            WriteLine("---------------------------*-----*----------------------------", ConsoleColor.Yellow);
            WriteLine("------------------------------*-------------------------------", ConsoleColor.Yellow);
            WriteLine("--------------------------------------------------------------", ConsoleColor.Yellow);
            WriteLine(string.Empty);
            ReadStartParams(args);
            WriteLine(string.Empty);
            WriteLine("Press enter to start.");
            Console.ReadLine();
            //WriteLine(GetCurrentDirectory());
            WriteLine("Reading files...");
            try
            {
                var colDataset = ReadColDataFolder();
                var outboundDataset = ReadOutboundDataFolder();
                WriteLine(string.Format("Col data: {0} rows.", colDataset.Tables[0].Rows.Count));
                WriteLine(string.Format("Outbound data: {0} rows.", outboundDataset.Tables[0].Rows.Count));
                if (colDataset.Tables[0].Rows.Count < 1 || outboundDataset.Tables[0].Rows.Count < 1)
                    throw new InvalidDataException("theres no COL or Outbound data");
                //report making method
                WriteLine("_________________________________________");
                WriteLine("Creating random contract report:", ConsoleColor.Cyan);
                CreateTakeRandomCallReport(colDataset, outboundDataset, randomReportTake, conversationLen); //take is configurable
                WriteLine("_________________________________________");
                WriteLine("Creating full report:", ConsoleColor.Cyan);
                CreateFullReport(colDataset, outboundDataset);
                //Test();
                WriteLine(string.Empty);
                WriteLine("...Done!");
                WriteLine(string.Empty);
                WriteLine("Press enter to end program.");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                WriteLine(ex.Message, ConsoleColor.Red);
                WriteLine(ex.StackTrace, ConsoleColor.Gray);
                if (ex.InnerException != null)
                {
                    WriteLine("Inner Ex:");
                    WriteLine(ex.InnerException.Message, ConsoleColor.Red);
                    WriteLine(ex.InnerException.StackTrace, ConsoleColor.Gray);
                }
                Console.ReadLine();
            }
        }
        private static void ReadStartParams(string[] args)
        {
            WriteLine("Reading argruments...");
            if(args.Length == 0)
            {
                WriteLine("No args -> use default values");
                return;
            }
            foreach (var param in args)
            {
                if(param.ToLower().Contains("callduration"))
                {
                    var split = param.Split(':');
                    if (split.Length < 2 || !int.TryParse(split.Last(), out int dur))
                    {
                        WriteLine(string.Format("Cant parse argrument: {0}", param), ConsoleColor.Red);
                        continue;
                    }
                    WriteLine(string.Format("Arg: Call duration: {0}", dur));
                    conversationLen = dur;
                    continue;
                }
                if (param.ToLower().Contains("take"))
                {
                    var split = param.Split(':');
                    if (split.Length < 2 || !int.TryParse(split.Last(), out int take))
                    {
                        WriteLine(string.Format("Cant parse argrument: {0}", param), ConsoleColor.Red);
                        continue;
                    }
                    WriteLine(string.Format("Arg: Take {0} random calls", take));
                    randomReportTake = take;
                    continue;
                }
                WriteLine(string.Format("Incorrect arg: {0} -> skip", param), ConsoleColor.Red);
            }
        }
        //col data is xlsx format, xls
        private static DataSet ReadColDataFolder()
        {
            var xlsxNames = GetFileNames(string.Format(@"{0}\{1}", GetCurrentDirectory(), ColFolder), "*.xlsx");
            var xlsNames = GetFileNames(string.Format(@"{0}\{1}", GetCurrentDirectory(), ColFolder), "*.xls");
            var fileNames = xlsxNames.Union(xlsNames).ToArray();
            WriteLine(string.Format("Found {0} files in {1}", fileNames.Length, ColFolder));
            if (fileNames.Length < 1)
                throw new FileNotFoundException(string.Format("No xlsx, xls file in {0} folder", ColFolder));
            //var set = OleReader.ReadExcelFile("sheet1", fileNames[0]);
            //for (int i = 1; i < fileNames.Length; i++)
            //    set.Merge(OleReader.ReadExcelFile("sheet1", fileNames[i]));
            //return set;
            var set = ExcelReader.GetDataTableFromExcel(fileNames[0]);
            for (int i = 1; i < fileNames.Length; i++)
                set.Merge(ExcelReader.GetDataTableFromExcel(fileNames[i]));
            return set;
        }

        //outbound data is CSV
        private static DataSet ReadOutboundDataFolder()
        {
            var fileNames = GetFileNames(string.Format(@"{0}\{1}", GetCurrentDirectory(), OutboundFolder), "*.csv");
            WriteLine(string.Format("Found {0} files in {1}", fileNames.Length, OutboundFolder));
            if (fileNames.Length < 1)
                throw new FileNotFoundException(string.Format("No csv file in {0} folder", OutboundFolder));
            var set = OleReader.ReadCSVFile(fileNames[0]);
            for (int i = 1; i < fileNames.Length; i++)
                set.Merge(OleReader.ReadCSVFile(fileNames[i]));
            return set;
        }

        private static string[] GetFileNames(string folderPath, string pattern)
        {
            return Directory.GetFiles(folderPath, pattern, SearchOption.TopDirectoryOnly);
        }

        //private static void Test()
        //{
        //    string fileName = string.Format("{0}/{1}", GetCurrentDirectory(), "merged_report.xlsx");
        //    var file = new FileInfo(fileName);
        //    if (File.Exists(fileName))
        //        File.Delete(fileName);
        //    using (var package = new ExcelPackage(file))
        //    {
        //        var worksheet = package.Workbook.Worksheets.Add("merged_report");
        //        worksheet.OutLineSummaryBelow = false;
        //        //cant set an array like excel interop
        //        //worksheet.SetValue("A1", new string[] { "cell1", "cell2", "cell3" });

        //        //ok
        //        //CollapseRow(2, 5, worksheet);
        //        //CollapseRow(15, 20, worksheet);

        //        //ok
        //        EPPlusHelper.SetRow(1, 2, new[] {"cell1", "cell2", "cell3"}, worksheet, Color.Yellow);

        //        package.Save();
        //    }

        //}

        //private static DataTable RandomOrderTable(DataTable table)
        //{
        //    string temp = "temp_sort";
        //    var cloneTable = table.Copy();
        //    cloneTable.Columns.Add(temp);
        //    var rnd = new Random(DateTime.Now.Millisecond);
        //    foreach (DataRow row in cloneTable.Rows)
        //    {
        //        row[temp] = rnd.Next(cloneTable.Rows.Count);
        //    }
        //    DataView dv = cloneTable.DefaultView;
        //    dv.Sort = temp;
        //    var randomTable = dv.ToTable();
        //    randomTable.Columns.Remove(temp);
        //    return randomTable;
        //}

        private static void UpdatePerfStat(int current, int max)
        {
            Console.Write("\rProcessed {0}/{1} rows", current, max);
        }

        private static Dictionary<string, List<int>> MapData(DataTable table, string columnToMapOn)
        {
            var dict = new Dictionary<string, List<int>>();
            int index = 0;
            foreach (DataRow row in table.Rows)
            {
                string cellContent = row[columnToMapOn].ToString();
                //if (string.IsNullOrEmpty(cellContent)) continue; //skip empty cell
                if (!dict.ContainsKey(cellContent))
                {
                    dict.Add(row[columnToMapOn].ToString(), new List<int> {index});
                    index++;
                    continue;
                }
                dict[row[columnToMapOn].ToString()].Add(index);
                index++;
            }
            return dict;
        }

        private static string[] GetHeaders(DataTable dt)
        {
            return dt.Columns.Cast<DataColumn>()
                .Select(x => x.ColumnName)
                .ToArray();
        }

        private static void WriteLine(string s, ConsoleColor? color = null)
        {
            if (color != null)
            {
                var prevColor = Console.ForegroundColor;
                Console.ForegroundColor = color ?? ConsoleColor.White;
                Console.WriteLine(s);
                Console.ForegroundColor = prevColor;
            }
            else
            {
                Console.WriteLine(s);
            }
        }

        private static string GetCurrentDirectory()
        {
            return Directory.GetCurrentDirectory();
        }
    }
}