using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace StatementTollWindow
{
    class RowsTrimmer
    {

        public static string TrimRows(string filepath, string lastCellToTrim, string sheetname, string newpath)
        {
            using (XLWorkbook wb = new XLWorkbook(filepath)) //подготовка файла отчета, посредством вырезания лишних строк
            {
                wb.TryGetWorksheet(sheetname, out IXLWorksheet ws);
                ws.Range($"A1:{lastCellToTrim}").Rows().Delete();
                wb.SaveAs(newpath);
                return newpath;
            }

        }
        public static string TrimRowsAndFixNaming(string filepath, string lastCellToTrim, string sheetname, string newpath)
        {
            using (XLWorkbook wb = new XLWorkbook(filepath)) //подготовка файла отчета, посредством вырезания лишних строк
            {
                wb.TryGetWorksheet(sheetname, out IXLWorksheet ws);
                ws.Range($"A1:{lastCellToTrim}").Rows().Delete();
                ws.Cell("R1").Value = "Sales";
                ws.Cell("X1").Value = "Total";

                wb.SaveAs(newpath);
                return newpath;
            }

        }



    }
}
