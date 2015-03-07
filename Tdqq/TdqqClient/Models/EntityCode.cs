using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TdqqClient.Models
{
    /// <summary>
    /// 实体与之相对应的代码
    /// </summary>
    public class EntityCode
    {
        public string Code { get; set; }
        public string Entity { get; set; }

        public EntityCode(string code, string entity)
        {
            Code = code;
            Entity = entity;
        }

        public EntityCode()
        {
            Code = Entity = string.Empty;
        }
    }
}
