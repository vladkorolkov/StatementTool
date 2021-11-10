using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StatementTollWindow
{
    class TableHandler
    {
        public static void resultTableHandler(OleDbDataReader dataReader, OleDbConnection connection, string artistName, string src, string filename)
        {

                XLWorkbook workbook = new XLWorkbook();
                var ws = workbook.Worksheets.Add("Example");
                int counter = 2;
                while (dataReader.Read())
                {
                    ws.Cell(counter, "A").Value = dataReader[0];
                    ws.Cell(counter, "B").Value = dataReader[1];
                    ws.Cell(counter, "C").Value = dataReader[2];
                    ws.Cell(counter, "D").Value = dataReader[3];
                    ws.Cell(counter, "E").Value = dataReader[4];
                    counter++;
                }

                ws.Cell("A1").Value = dataReader.GetName(0);
                ws.Cell("B1").Value = dataReader.GetName(1);
                ws.Cell("C1").Value = dataReader.GetName(2);
                ws.Cell("D1").Value = dataReader.GetName(3);
                ws.Cell("E1").Value = dataReader.GetName(4);
                
                ws.Cell("G1").Value = "TotalCash";
                ws.Cell("H1").Value = "TotalListens";


                ws.Range("A1:H1").Rows().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Range("A1:H1").Rows().Style.Font.Bold = true;
                ws.Column("B").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                dataReader.Close();

                if (src == "NDA")
                {
                    OleDbCommand sumCommand = new OleDbCommand($"SELECT SUM([Total]),SUM([Sales]) FROM [1 Детализированный отчет$] WHERE [Исполнитель] = '{artistName}'");
                    sumCommand.Connection = connection;
                    dataReader = sumCommand.ExecuteReader();

                    while (dataReader.Read())
                    {
                        var result = Convert.ToSingle(dataReader[0]); // конвертация для того чтобы округление сработало                   
                        ws.Cell("G2").Value = Math.Round(result, 3);
                        ws.Cell("H2").Value = dataReader[1];
                    }
                    dataReader.Close();
                }
               
              if (src == "FT")
                {
                    OleDbCommand sumCommand = new OleDbCommand($"SELECT SUM([Sale count]),SUM([Amount]) FROM [Sheet1$] WHERE [Artist name] = '{artistName}'");
                    sumCommand.Connection = connection;
                    dataReader = sumCommand.ExecuteReader();

                    while (dataReader.Read())
                    {
                        var result = Convert.ToSingle(dataReader[0]); // конвертация для того чтобы округление сработало                   
                        ws.Cell("G2").Value = Math.Round(result, 3);
                        ws.Cell("H2").Value = dataReader[1];
                    }
                    dataReader.Close();
                }
            //workbook.SaveAs(@"C:\Users\vladi\source\repos\HelloWorld.xlsx");                
            
            workbook.SaveAs(filename);
            // workbook.Save();
           
           

        }

       

        
    }
}
