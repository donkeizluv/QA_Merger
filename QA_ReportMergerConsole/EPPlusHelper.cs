using System.Drawing;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace QA_ReportMergerConsole
{
    public static class EPPlusHelper
    {
        public static ExcelRange SetRow(int row, int col, object[] contentString, ExcelWorksheet ws, Color? color = null)
        {
            var result = contentString.Where(x => x != null)
                .Select(x => x.ToString())
                .ToArray();
            return SetRow(row, col, result, ws, color);
        }

        public static ExcelRange SetRow(int row, int col, string[] contentString, ExcelWorksheet ws, Color? color = null)
        {
            int tempCol = col;
            foreach (string s in contentString)
            {
                ws.Cells[row, tempCol].Value = s;
                if (color != null)
                {
                    ws.Cells[row, tempCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[row, tempCol].Style.Fill.BackgroundColor.SetColor(color ?? Color.White);
                }
                tempCol++;
            }
            return ws.Cells[row, col, row, tempCol - 1];
        }

        public static void CollapseRow(int from, int to, ExcelWorksheet ws)
        {
            for (int i = from; i < to; i++)
            {
                ws.Row(i).OutlineLevel = 1;
                ws.Row(i).Collapsed = true;
            }
        }
    }
}