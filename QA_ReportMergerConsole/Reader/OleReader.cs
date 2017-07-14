using System.Data;
using System.Data.OleDb;
using System.IO;

namespace QA_ReportMergerConsole.Reader
{
    public static class OleReader
    {
        //enum ExcelFileMode {XLSX, XLS};
        public static DataSet ReadExcelFile(string sheetName, string path)
        {
            using (var conn = new OleDbConnection())
            {
                var dt = new DataTable();
                string Import_FileName = path;
                string fileExtension = Path.GetExtension(Import_FileName);
                //var mode = ExcelFileMode.XLSX;
                if (fileExtension == ".xls")
                {
                    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Import_FileName + ";" +
                                            "Extended Properties='Excel 8.0;HDR=YES;'";
                    //mode = ExcelFileMode.XLS;
                }

                if (fileExtension == ".xlsx")
                {
                    conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Import_FileName + ";" +
                                            "Extended Properties='Excel 12.0 Xml;HDR=YES;'";
                }
                //conn.Open();
                using (var comm = new OleDbCommand())
                {
                    comm.CommandText = "Select * from [" + sheetName + "$]";
                    //if (mode == ExcelFileMode.XLS)
                    //{
                    //    comm.CommandText = "Select * from [" + sheetName + "]";
                    //}

                    comm.Connection = conn;
                    //get all sheets name
                    //var dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    using (var da = new OleDbDataAdapter())
                    {
                        var ds = new DataSet();
                        da.SelectCommand = comm;
                        da.Fill(ds);
                        return ds;
                    }
                }
            }
        }

        public static DataSet ReadCSVFile(string path)
        {
            //var filename = @"c:\work\test.csv";
            string connString = string.Format(
                @"Provider=Microsoft.Jet.OleDb.4.0; Data Source={0};Extended Properties=""Text;HDR=YES;FMT=Delimited""",
                Path.GetDirectoryName(path)
            );
            using (var conn = new OleDbConnection(connString))
            {
                conn.Open();
                string query = "SELECT * FROM [" + Path.GetFileName(path) + "]";
                using (var adapter = new OleDbDataAdapter(query, conn))
                {
                    var ds = new DataSet();
                    adapter.Fill(ds);
                    return ds;
                }
            }
        }
    }
}