using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ESRI.ArcGIS.Geodatabase;

namespace TdqqClient.Models
{
    public class TdqqField
    {
         public esriFieldType FieldType { get; set; }
        public string FieldName { get; set; }
        public int Length { get; set; }

        public TdqqField(esriFieldType fieldType, string fieldName, int length)
        {
            FieldType = fieldType;
            FieldName = fieldName;
            Length = length;
        }
    }
}
