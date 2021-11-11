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
            ($"SELECT [Artist Name],[Album Name],[Platform],[Sale count],[Amount],[Date] FROM [Sheet1$] WHERE [Artist Name] = '{artistName}'");
            command.Connection = connection;
            command.CommandType = CommandType.Text;
            return command;
        }

        public static OleDbCommand sqlNda(OleDbConnection connection, string artistName)
        {
            OleDbCommand command = new OleDbCommand
      ($"SELECT [Исполнитель],[Название альбома],[Площадка],[Sales],[Total],[Период использования контента] FROM [1 Детализированный отчет$] WHERE [Исполнитель] = '{artistName}'");
            command.Connection = connection;
            command.CommandType = CommandType.Text;
            return command;
        }

     
    }
}
