using System.Data.OleDb;
using TdqqClient.Services.Database;

namespace TdqqClient.Services.Import
{
    /// <summary>
    /// 导入至数据库
    /// </summary>
    class ImportToDb:IImport
    {
        private readonly  IDatabaseService _pDatabaseService;

        public ImportToDb(IDatabaseService pDatabaseService)
        {
            _pDatabaseService = pDatabaseService;
        }

        /// <summary>
        /// 删除表
        /// </summary>
        /// <param name="tableName">表的名称</param>
        /// <returns>是否成功</returns>
        public bool DeleteTable(string tableName)
        {
            var sqlString = string.Format("delete from {0}", tableName);
            var ret = _pDatabaseService.Execute(sqlString);
            return ret == -1 ? false : true;
        }

        /// <summary>
        /// 插入一行数据
        /// </summary>
        /// <param name="insertExpression">插入一行数据的sql语句</param>
        /// <returns>是否插入一行数据成功</returns>
        public bool InsertRow(string insertExpression)
        {
            var ret = _pDatabaseService.Execute(insertExpression);
            return ret == -1 ? false : true;
        }

        public bool UpdateColumn(string sqlString)
        {
            var ret = _pDatabaseService.Execute(sqlString);
            return ret == -1 ? false : true;
        }

        public System.Data.DataTable Query(string sqlString)
        {
            return _pDatabaseService.Query(sqlString);
        }

        public OleDbConnection DbConnection()
        {
            return _pDatabaseService.Connnection();
        }
    }
}
