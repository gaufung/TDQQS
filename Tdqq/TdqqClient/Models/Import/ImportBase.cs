using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TdqqClient.Services.Database;

namespace TdqqClient.Models.Import
{
    class ImportBase
    {
        /// <summary>
        /// 基础数据库的路径
        /// </summary>
        protected string BasicDatabase { get; set; }
        public ImportBase(string basicDatabase)
        {
            BasicDatabase = basicDatabase;
        }
        public ImportBase()
        {
            BasicDatabase = string.Empty;
        }
        protected bool DeleteTable(string tableName)
        {
            var sqlString = string.Format("delete from {0}", tableName);
            var accessFactory = new MsAccessDatabase(BasicDatabase);
            var ret = accessFactory.Execute(sqlString);
            return ret == -1 ? false : true;
        }
        protected bool InsertRow(string insertExpression)
        {

            var accessfactory = new MsAccessDatabase(BasicDatabase);
            var ret = accessfactory.Execute(insertExpression);
            return ret == -1 ? false : true;
        }
        /// <summary>
        /// 导入信息
        /// </summary>
        /// <returns></returns>
        public virtual void Import()
        {
            
        }

        
        
    }
}
