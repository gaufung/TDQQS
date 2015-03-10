using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TdqqClient.Models
{
    /// <summary>
    /// 为了地块编码而排序设计的实体，采用的是泛型设计
    /// </summary>
    /// <typeparam name="T">实体的类型</typeparam>
    public class SortEnity<T>
    {
        public T Id { get; set; }
        public double Xcor { get; set; }
        public double Ycor { get; set; }
    }
}
