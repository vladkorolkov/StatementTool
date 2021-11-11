using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data.OleDb;
using Microsoft.Win32;

namespace StatementTollWindow
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
      
        string artistName { get; set; }
        string filepath { get; set; }
        string workingfile;
        string sourceType = "";

        public MainWindow()
        {
            InitializeComponent();
            this.Closed += MainWindow_Closed;
            this.Closing += MainWindow_Closing;
          
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog opfl = new Microsoft.Win32.OpenFileDialog();
            opfl.FileName = "Document";
            opfl.DefaultExt = ".xlsx";
            opfl.Filter = "Excel documents (.xlsx)|*.xlsx";
            Nullable<bool> result = opfl.ShowDialog();
            if (result == true)
            {
                filepath = opfl.FileName;               
            }
       
        }

        private void MainWindow_Closing (object sender, System.ComponentModel.CancelEventArgs e)
        {
            string msg = "Уверены что хотите закрыть окно и завершить программу?";
            MessageBoxResult result = MessageBox.Show(msg, "Myapp", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if(result == MessageBoxResult.No)
            {
                e.Cancel = true;
            }
        }

        private void MainWindow_Closed (object sender, EventArgs e)
        {
            //MessageBox.Show("Всего хорошего!");
        }

        private void FT_Checked(object sender, RoutedEventArgs e)
        {
            RadioButton pressed = (RadioButton)sender;
            //MessageBox.Show(pressed.Content.ToString());
            sourceType = "FT";
            workingfile = RowsTrimmer.TrimRows(filepath, "L3", "Sheet1", @"C:\Users\vladi\source\repos\trimmedReport.xlsx");
            //string constring = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={workingfile};Extended Properties=\'Excel 12.0 Xml;HDR=YES;IMEX=2\'";
            //OleDbConnection connection = new OleDbConnection(constring);
            //connection.Open();
            //OleDbDataReader dataReader = Commands.sqlFreshTunes(connection, artistName).ExecuteReader();
            //TableHandler.FreshTunesHandler(dataReader, connection, artistName);
            //dataReader.Close();
        }

        private void NDA_Checked(object sender, RoutedEventArgs e) 
        {
            workingfile = RowsTrimmer.TrimRowsAndFixNaming(filepath, "Y5", "1 Детализированный отчет", @"C:\Users\vladi\source\repos\trimmedReport.xlsx");
            sourceType = "NDA";
            //string constring = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={workingfile};Extended Properties=\'Excel 12.0 Xml;HDR=YES;IMEX=2\'";
            //OleDbConnection connection = new OleDbConnection(constring);
            //connection.Open();
            //OleDbDataReader datareader = Commands.sqlNda(connection, artistName).ExecuteReader();
            //TableHandler.FreshTunesHandler(datareader, connection, artistName);
            //datareader.Close();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            artistName = artistBox.Text;
            MessageBox.Show("Выбранное имя артиста: " + artistName);
        }

        private void MakeReport_Click(object sender, RoutedEventArgs e)
        {

            string constring = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={workingfile};Extended Properties=\'Excel 12.0 Xml;HDR=YES;IMEX=2\'";
            OleDbConnection connection = new OleDbConnection(constring);
            connection.Open();
            if (sourceType == "FT")
            {
                OleDbDataReader dataReader = Commands.sqlFreshTunes(connection, artistName).ExecuteReader();

                SaveFileDialog sfd = new SaveFileDialog();
                sfd.DefaultExt = ".xlsx";
                sfd.Filter = "Excel files(*.xlsx)|*.xlsx|All Files(*.*)|*.*";
                if (sfd.ShowDialog() == true)
                {                   
                    var filename = sfd.FileName;
                    TableHandler.resultTableHandler(dataReader, connection, artistName, sourceType,filename);
                }
                
                             
                dataReader.Close();
            }
            if(sourceType == "NDA")
            {
                OleDbDataReader dataReader = Commands.sqlNda(connection, artistName).ExecuteReader();

                SaveFileDialog sfd = new SaveFileDialog();
                sfd.DefaultExt = ".xlsx";
                sfd.Filter = "Excel files(*.xlsx)|*.xlsx|All Files(*.*)|*.*";
                if (sfd.ShowDialog() == true)
                {        
                    var filename = sfd.FileName;
                    TableHandler.resultTableHandler(dataReader, connection, artistName, sourceType, filename);
                }
                
                
                dataReader.Close();
            }

         


        }
       
       
    }
}
