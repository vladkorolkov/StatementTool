using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;

namespace StatementTollWindow
{
    class Commands
    {
        public static OleDbCommand sqlFreshTunes(OleDbConnection connection, string artistName)
        {
            OleDbCommand command = new OleDbCommand
            ($"SELECT [Artist Name],[Platform],[Amount],[Sale count],[Date] FROM [Sheet1$] WHERE [Artist Name] = '{artistName}'");
            command.Connection = connection;
            command.CommandType = CommandType.Text;
            return command;
        }

        public static OleDbCommand sqlNda(OleDbConnection connection, string artistName)
        {
            OleDbCommand command = new OleDbCommand
      ($"SELECT [Исполнитель],[Период использования контента],[Площадка],[Sales],[Total] FROM [1 Детализированный отчет$] WHERE [Исполнитель] = '{artistName}'");
            command.Connection = connection;
            command.CommandType = CommandType.Text;
            return command;
        }

     
    }
}
