

using System.Data.SqlClient;

namespace TdqqClient.Services.Import
{
    interface IImport
    {
        bool DeleteTable(string tableName);
        bool InsertRow(string insertExpression);

        bool UpdateColumn(string sqlString);

        System.Data.DataTable Query(string sqlString);

        
    }
}
