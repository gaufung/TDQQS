using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ESRI.ArcGIS.Geodatabase;
using TdqqClient.Services.AE;
using TdqqClient.Services.Check;

namespace TdqqClient.Models.Edit
{
    /// <summary>
    /// 编辑字段子类
    /// </summary>
    class EditFields:EditModel
    {
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="personDatabase">个人地理数据库</param>
        /// <param name="selectFeauture">选择的要素类</param>
        /// <param name="basicDatabase">基础数据库</param>
        public EditFields(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }

        public override void Edit(object parameter)
        {
            IAeFactory pAeFactory = new PersonalGeoDatabase(PersonDatabase);
            IFeatureClass pFeatureClass = pAeFactory.OpenFeatureClasss(SelectFeature);
            if (!pFeatureClass.FieldExistCheck("YHTMJ", "CBFMC"))
            {
                System.Windows.Forms.MessageBox.Show(null,
                    "缺少部分字段", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (!CheckFieldType(pFeatureClass))
            {
                System.Windows.Forms.MessageBox.Show(null,
                    "字段类型不符合要求", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (!DeleteFields())
            {
                System.Windows.Forms.MessageBox.Show(null,
                    "删除字段失败", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (!AddFields())
            {
                System.Windows.Forms.MessageBox.Show(null,
                    "添加字段失败", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            System.Windows.Forms.MessageBox.Show(null,
                    "编辑字段成功", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);    
        }

        /// <summary>
        /// 检查必备字段是否存在
        /// </summary>
        /// <returns></returns>
        private bool CheckFieldType(IFeatureClass pFeatureClass)
        {
            bool flag = true;
            //原合同面积
            if (!pFeatureClass.FieldTypeCheck("YHTMJ", esriFieldType.esriFieldTypeDouble) &&
                !pFeatureClass.FieldTypeCheck("YHTMJ", esriFieldType.esriFieldTypeSingle))
            {
                flag = false;
            }
            //承包方名称
            if (!pFeatureClass.FieldTypeCheck("CBFMC", esriFieldType.esriFieldTypeString))
            {
                flag = false;
            }
            //地块名称
            if (!pFeatureClass.FieldTypeCheck("DKMC", esriFieldType.esriFieldTypeString))
            {
                flag = false;
            }
            return flag;
        }   

        /// <summary>
        /// 删除无效字段
        /// </summary>
        /// <returns></returns>
        private bool DeleteFields()
        {
            var pAeFactory = new PersonalGeoDatabase(PersonDatabase);      
            IFeatureClass pFeatureClass = pAeFactory.OpenFeatureClasss(SelectFeature);
            bool flag;
            try
            {
                var toDeleteFields = GetToDeleteFields(pFeatureClass);
                pAeFactory.DeleteFields(pFeatureClass, toDeleteFields);
                pAeFactory.ReleaseFeautureClass(pFeatureClass);
                flag = true;
            }
            catch (Exception)
            {
                flag = false;
            }
            finally
            {
                pAeFactory.ReleaseFeautureClass(pFeatureClass);
            }
            return flag;
        }  


        /// <summary>
       /// 获取要删除的字段
       /// </summary>
       /// <param name="pFeatureClass">地块要素类</param>
       /// <returns>删除的集合</returns>
        private List<IField> GetToDeleteFields(IFeatureClass pFeatureClass)
        {
            var toDeleteFields = new List<IField>();
            for (int i = 0; i < pFeatureClass.Fields.FieldCount; i++)
            {
                var fieldName = pFeatureClass.Fields.Field[i].Name.Trim().ToLower();
                //要素自带字段
                if (fieldName == "objectid" || fieldName == "shape" || fieldName == "shape_length" || fieldName == "shape_area" || fieldName == "bl" || fieldName == "clipt") continue;
                //字符串类型
                if (fieldName == "cbfmc" || fieldName == "dkmc" || fieldName == "dkbm" || fieldName == "dkdz" || fieldName == "dknz" ||
                    fieldName == "dkbz" || fieldName == "dkxz" || fieldName == "dkbzxx" || fieldName == "zjrxm" || fieldName == "fbfbm" || fieldName == "cbfbm"
                    || fieldName == "lzhtbm" || fieldName == "cbhtbm" || fieldName == "syqxz" || fieldName == "dklb" || fieldName == "tdlylx" || fieldName == "dldj"
                    || fieldName == "tdyt" || fieldName == "sfjbnt" || fieldName == "cbjyqqdfs" || fieldName == "ysdm")
                {
                    if (pFeatureClass.Fields.Field[i].Type == esriFieldType.esriFieldTypeString) continue;
                }
                //double数字类型
                if (fieldName == "yhtmj" || fieldName == "htmj" || fieldName == "scmj")
                {
                    if (pFeatureClass.Fields.Field[i].Type == esriFieldType.esriFieldTypeDouble || pFeatureClass.Fields.Field[i].Type == esriFieldType.esriFieldTypeSingle)
                    {
                        continue;
                    }
                }
                toDeleteFields.Add(pFeatureClass.Fields.Field[i]);
            }
            return toDeleteFields;
        }   

         /// <summary>
        /// 增加字段
        /// </summary>
        /// <returns></returns>
        private bool AddFields()
        {
            try
            {
                IAeFactory pAeFactory = new PersonalGeoDatabase(PersonDatabase);
                var toAddFields = GetToAddFields();
                foreach (var tdqqField in toAddFields)
                {
                    pAeFactory.AddField(SelectFeature, tdqqField.FieldName, tdqqField.Length, tdqqField.FieldType);
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private IEnumerable<TdqqField> GetToAddFields()
        {
            //string 类型的
            List<TdqqField> listFields = new List<TdqqField>();

            listFields.Add(new TdqqField(esriFieldType.esriFieldTypeString, "DKBM", 19));
            listFields.Add(new TdqqField(esriFieldType.esriFieldTypeString, "DKMC", 50));
            listFields.Add(new TdqqField(esriFieldType.esriFieldTypeString, "DKDZ", 50));
            listFields.Add(new TdqqField(esriFieldType.esriFieldTypeString, "DKNZ", 50));
            listFields.Add(new TdqqField(esriFieldType.esriFieldTypeString, "DKXZ", 50));
            listFields.Add(new TdqqField(esriFieldType.esriFieldTypeString, "DKBZ", 50));
            listFields.Add(new TdqqField(esriFieldType.esriFieldTypeString, "DKBZXX", 300));
            listFields.Add(new TdqqField(esriFieldType.esriFieldTypeString, "ZJRXM", 100));
            listFields.Add(new TdqqField(esriFieldType.esriFieldTypeString, "FBFBM", 14));
            listFields.Add(new TdqqField(esriFieldType.esriFieldTypeString, "CBFBM", 18));
            listFields.Add(new TdqqField(esriFieldType.esriFieldTypeString, "CBJYQZBM", 19));
            listFields.Add(new TdqqField(esriFieldType.esriFieldTypeString, "LZHTBM", 20));
            listFields.Add(new TdqqField(esriFieldType.esriFieldTypeString, "CBHTBM", 18));
            listFields.Add(new TdqqField(esriFieldType.esriFieldTypeString, "SYQXZ", 2));
            listFields.Add(new TdqqField(esriFieldType.esriFieldTypeString, "DKLB", 2));
            listFields.Add(new TdqqField(esriFieldType.esriFieldTypeString, "TDLYLX", 3));
            listFields.Add(new TdqqField(esriFieldType.esriFieldTypeString, "DLDJ", 2));
            listFields.Add(new TdqqField(esriFieldType.esriFieldTypeString, "TDYT", 1));
            listFields.Add(new TdqqField(esriFieldType.esriFieldTypeString, "SFJBNT", 1));
            listFields.Add(new TdqqField(esriFieldType.esriFieldTypeString, "CBJYQQDFS", 3));
            listFields.Add(new TdqqField(esriFieldType.esriFieldTypeString, "YSDM", 6));
            //double 类型的
            listFields.Add(new TdqqField(esriFieldType.esriFieldTypeDouble, "HTMJ", 15));
            listFields.Add(new TdqqField(esriFieldType.esriFieldTypeDouble, "SCMJ", 15));
            //int 类型
            listFields.Add(new TdqqField(esriFieldType.esriFieldTypeInteger, "BSM", 10));
            return listFields;
        }
       
    }
}
