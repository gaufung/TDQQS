using System;
using System.Data;
using System.Data.OleDb;
namespace TdqqClient.Services.Database
{
    /// <summary>
    /// Access数据库操作
    /// </summary>
    class MsAccessDatabase:IDatabaseService
    {
        private readonly  string _basicDatabase;
        public MsAccessDatabase(string basicDatabase)
        {
            _basicDatabase = basicDatabase;
        }
        private OleDbConnection GetConnection()
        {
            try
            {
                string connnectString = "Provider=Microsoft.Jet.OLEDB.4.0;" + "data source=" + _basicDatabase;
                var con = new OleDbConnection(connnectString);
                return con;
            }
            catch (Exception)
            {
                return null;
            }
        }
        public System.Data.DataTable Query(string sqlString)
        {
            var con = GetConnection();
            if (con == null) return null;
            con.Open();
            var dt = new DataTable();
            if (con.State != ConnectionState.Open)return null;
            try
            {
                var adapter = new OleDbDataAdapter(sqlString,con);              
                adapter.Fill(dt);
            }
            catch (Exception)
            {
                dt = null;
            }
            finally
            {
                con.Close();               
            }
            return dt;
        }
        public int Execute(string sqlString)
        {
            const int errorState = -1;
            var con = GetConnection();
            if (con == null) return errorState;
            con.Open();
            if (con.State != ConnectionState.Open) return errorState;
            int ret;
            try
            {
                var cmd = new OleDbCommand(sqlString, con);
                ret = cmd.ExecuteNonQuery();
            }
            catch (Exception)
            {
               
                ret = errorState;
            }
            finally
            {
                con.Close();
            }
            return ret;
        }
        public object ExecuteScalar(string sqlString)
        {
            var con = GetConnection();
            //数据库连接对象为空
            if (con == null)return null;
            con.Open();
            if (con.State != ConnectionState.Open) return null;
            var ret = new object();
            try
            {
                var cmd = new OleDbCommand(sqlString, con);
                ret = cmd.ExecuteScalar();
            }
            catch (Exception)
            {
                ret = null;
            }
            finally
            {
                con.Close();
            }
            return ret;
        }
    }
}
