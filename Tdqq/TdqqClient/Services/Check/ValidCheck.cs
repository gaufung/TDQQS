using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Geometry;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using TdqqClient.Services.AE;
using TdqqClient.Services.Database;
using IRow = NPOI.SS.UserModel.IRow;

namespace TdqqClient.Services.Check
{
    /// <summary>
    /// 有效性检查
    /// </summary>
    class ValidCheck
    {
        #region 家庭成员信息表检查
        /// <summary>
        /// 家庭成员信息表的列表是否按照标准数据
        /// </summary>
        /// <param name="excelPath">excel文件地址</param>
        /// <returns>是否满足条件</returns>
        public static bool ExcelColumnSorted(string excelPath)
        {
            try
            {
                using (var fileStream = new FileStream(excelPath, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook = new HSSFWorkbook(fileStream);
                    ISheet sheet = workbook.GetSheetAt(0);
                    IRow row = sheet.GetRow(0);
                    bool flag;
                    if (row.GetCell(0).ToString().Trim() != "CBFBM" || row.GetCell(1).ToString().Trim() != "CYXB"
                        || row.GetCell(2).ToString().Trim() != "CYXM" || row.GetCell(3).ToString().Trim() != "CYZJLX" ||
                        row.GetCell(4).ToString().Trim() != "CYZJHM" || row.GetCell(5).ToString().Trim() != "CYBZ"
                        || row.GetCell(6).ToString().Trim() != "YHZGX" || row.GetCell(7).ToString().Trim() != "CYSZC" ||
                        row.GetCell(8).ToString().Trim() != "YZBM" || row.GetCell(9).ToString().Trim() != "SFGYR" ||
                        row.GetCell(10).ToString().Trim() != "LXDH") flag = false;
                    else
                    {
                        flag = true;
                    }
                    fileStream.Close();
                    return flag;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// 检查Excel中一行数据是否符合规范
        /// </summary>
        /// <param name="row">一行的对象</param>
        /// <param name="errorInfo">错误信息</param>
        /// <returns>返回是否满足要求</returns>
        public static bool ExcelRowCheck(IRow row, ref string errorInfo)
        {
            //承包方编码
            var cell = row.GetCell(0);
            if (cell == null || cell.ToString().Trim().Length != 18)
            {
                errorInfo = "承包方编码错误";
                return false;
            }
            //成员性别
            cell = row.GetCell(1);
            if (cell == null || cell.ToString().Length != 1)
            {
                errorInfo = "成员性别错误";
                return false;
            }
            //成员名称
            cell = row.GetCell(2);
            if (cell == null || cell.ToString().Trim().Length >= 50)
            {
                errorInfo = "成员名称错误";
                return false;
            }
            //成员证件类型
            cell = row.GetCell(3);
            if (cell != null && cell.ToString().Trim().Length != 1)
            {
                errorInfo = "证件类型代码错误";
                return false;
            }
            //成员证件号码
            cell = row.GetCell(4);
            if (cell != null && cell.ToString().Trim().Length > 20)
            {
                errorInfo = "证件号码大于20位";
                return false;
            }
            //与户主关系
            cell = row.GetCell(6);
            if (cell == null || cell.ToString().Trim().Length != 2)
            {
                errorInfo = "与户主关系代码错误";
                return false;
            }
            //邮政编码
            cell = row.GetCell(8);
            if (cell != null && cell.ToString().Trim().Length > 6)
            {
                var leng = cell.ToString().Trim().Length;
                errorInfo = "邮政编码代码错误";
                return false;
            }
            //联系电话
            cell = row.GetCell(10);
            if (cell != null && cell.ToString().Trim().Length > 20)
            {
                errorInfo = "联系电话长度大于20";
                return false;
            }
            return true;
        } 
        #endregion

        #region 地理数据库检查
        /// <summary>
        /// 检查字段是否为空
        /// </summary>
        /// <param name="personDatabase"></param>
        /// <param name="selectFeature"></param>
        /// <param name="fieldName"></param>
        /// <returns></returns>
        public static bool PersonDatabaseNullField(string personDatabase, string selectFeature, string fieldName)
        {
            IDatabaseService pDatabaseService = new MsAccessDatabase(personDatabase);
            var sqlString = string.Format("Select Count(*) From {0} Where {1} is null", selectFeature, fieldName);
            var res = pDatabaseService.ExecuteScalar(sqlString);
            if (res == null) return false;
            return (int)res > 0 ? false : true;
        }

        /// <summary>
        /// 检查某个字段是否满足是否规定的格式
        /// </summary>
        /// <param name="persondDatabase">个人地理数据库</param>
        /// <param name="selectFeaure">选择的要素类</param>
        /// <param name="fieldName">字段名称</param>
        /// <param name="targetFieldType">检查的类型</param>
        /// <returns>字段类型是否一致</returns>
        public static bool FieldTypeCheck(string persondDatabase, string selectFeaure, string fieldName,
            esriFieldType targetFieldType)
        {
            IAeFactory pAeFactory = new PersonalGeoDatabase(persondDatabase);
            IFeatureClass pFeatureClass = pAeFactory.OpenFeatureClasss(selectFeaure);
            bool flag;
            if (pFeatureClass.Fields.FindField(fieldName) == -1) flag = true;
            else
            {
                var type = pFeatureClass.Fields.Field[pFeatureClass.Fields.FindField(fieldName)].Type;
                flag = type == targetFieldType;
            }
            pAeFactory.ReleaseFeautureClass(pFeatureClass);
            return flag;
        }

        /// <summary>
        /// 检查字段是否存在
        /// </summary>
        /// <param name="personDatabase">个人地理数据库</param>
        /// <param name="selectFeaure">选择的要素类</param>
        /// <param name="fields">字段集合</param>
        /// <returns>只要有一个不存在则返回flag</returns>
        public static bool FieldExistCheck(string personDatabase, string selectFeaure, params string[] fields)
        {
            IAeFactory pAeFactory = new PersonalGeoDatabase(personDatabase);
            IFeatureClass pFeatureClass = pAeFactory.OpenFeatureClasss(selectFeaure);
            bool flag = true;
            foreach (var field in fields)
            {
                if (pFeatureClass.FindField(field) == -1)
                {
                    flag = false;
                    break;
                }
            }
            pAeFactory.ReleaseFeautureClass(pFeatureClass);
            return flag;
        }

        /// <summary>
        /// 检查要素类是否为地块要素
        /// </summary>
        /// <param name="personDatabase">个人地理数据库</param>
        /// <param name="selectFeature">选择的要素类</param>
        /// <param name="toCheckGeometryType">要检查的要素类型</param>
        /// <returns>是否满足要求</returns>
        public static bool CheckFeatureClassType(string personDatabase, string selectFeature, esriGeometryType toCheckGeometryType)
        {
            IAeFactory pAeFactory = new PersonalGeoDatabase(personDatabase);
            var pFeatureClass = pAeFactory.OpenFeatureClasss(selectFeature);
            return pFeatureClass == null ? false : toCheckGeometryType == pFeatureClass.ShapeType;
        }

        /// <summary>
        /// 判断要善于类是否存在
        /// </summary>
        /// <param name="personDatabase">个人地理数据库</param>
        /// <param name="feaureClassName">检查的要素类</param>
        /// <returns>是否存在</returns>
        public static bool IsExist(string personDatabase, string feaureClassName)
        {
            IAeFactory pAeFactory=new PersonalGeoDatabase(personDatabase);
            IFeatureWorkspace workspace = pAeFactory.OpenWorkspace();
            IEnumDataset dataset = (workspace as IWorkspace).get_Datasets(esriDatasetType.esriDTAny);
            IDataset tmp = null;
            while ((tmp = dataset.Next()) != null)
            {
                if (tmp.Name == feaureClassName)
                {
                    break;
                }
            }
            if (tmp != null)
            {
                return true;
            }
            return false;
        }
        #endregion



    }
}
