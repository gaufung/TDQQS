using System.Data;
using System.Data.OleDb;

namespace TdqqClient.Services.Database
{
    /// <summary>
    /// 定义数据库操作借口
    /// </summary>
    interface IDatabaseService
    {
        /// <summary>
        /// 查询操作
        /// </summary>
        /// <param name="sqlString">SQL执行语句</param>
        /// <returns>数据表</returns>
        DataTable Query(string sqlString);

        /// <summary>
        /// 执行操作
        /// </summary>
        /// <param name="sqlString">SQL执行语句</param>
        /// <returns>执行操作的数目</returns>
        int Execute(string sqlString);

        /// <summary>
        /// 返回单个值的执行方法
        /// </summary>
        /// <param name="sqlString"></param>
        /// <returns></returns>
        object ExecuteScalar(string sqlString);

        OleDbConnection Connnection();
    }
}
