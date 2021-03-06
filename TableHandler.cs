using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

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
                    ws.Cell(counter, "E").Value = Math.Round(Convert.ToSingle(dataReader[4]),3);
                    ws.Cell(counter, "F").Value = dataReader[5];
                    
                    counter++;
                }

                ws.Cell("A1").Value = "Название проекта";
                ws.Cell("B1").Value = "Релиз";
                ws.Cell("C1").Value = "Платформа";
                ws.Cell("D1").Value = "Прослушивания";
                ws.Cell("E1").Value = "Сумма";
                ws.Cell("F1").Value=  "Дата";
                ws.Cell("G1").Value = "Всего прослушиваний";
                ws.Cell("H1").Value = "Всего сумма";

                ws.Range("A1:H1").Rows().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Range("A1:H1").Rows().Style.Font.Bold = true;
                ws.Columns(1, 8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
  
                dataReader.Close();
            try
            {
                if (src == "NDA")
                {
                    OleDbCommand sumCommand = new OleDbCommand($"SELECT SUM([Sales]),SUM([Total]) FROM [1 Детализированный отчет$] WHERE [Исполнитель] = '{artistName}'");
                    sumCommand.Connection = connection;
                    dataReader = sumCommand.ExecuteReader();

                    while (dataReader.Read())
                    {
                        var resultG = Convert.ToSingle(dataReader[0]); // конвертация для того чтобы округление сработало                   
                        ws.Cell("G2").Value = Math.Round(resultG, 2);
                        var resultH = Convert.ToSingle(dataReader[1]);
                        ws.Cell("H2").Value = Math.Round(resultH, 2);
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
                        var resultG = Convert.ToSingle(dataReader[0]); // конвертация для того чтобы округление сработало                   
                        ws.Cell("G2").Value = Math.Round(resultG, 2);
                        var resultH = Convert.ToSingle(dataReader[1]);
                        ws.Cell("H2").Value = Math.Round(resultH, 2);

                    }
                    dataReader.Close();
                }
                              
            }

            catch (Exception ex)
            {
                MessageBox.Show($"Артиста с именем {artistName} не существует.");
            }
            workbook.SaveAs(filename);
           
           
           

        }

       

        
    }
}
