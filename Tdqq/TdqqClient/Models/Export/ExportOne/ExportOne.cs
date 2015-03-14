using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TdqqClient.Models.Export.ExportOne
{
    /// <summary>
    /// 一个村只要一个一张表的输出
    /// </summary>
    class ExportOne:ExportBase
    {
        public ExportOne(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }

        /// <summary>
        /// 导出结果，作为委托的
        /// </summary>
        /// <param name="parameter"></param>
        /// <returns></returns>
        public virtual void Export(object parameter)
        {
                           
        }
    }
}
