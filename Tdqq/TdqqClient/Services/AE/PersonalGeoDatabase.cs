using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using ESRI.ArcGIS.Geodatabase;

namespace TdqqClient.Services.AE
{
    class PersonalGeoDatabase:IAeFactory
    {
        /// <summary>
        /// 个人地理书数据库文件的路径
        /// </summary>
        private readonly string _personalDatabsePath;
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="personalDatabase">路径</param>
        public PersonalGeoDatabase(string personalDatabase)
        {
            _personalDatabsePath = personalDatabase;
        }

        public ESRI.ArcGIS.Geodatabase.IFeatureWorkspace OpenWorkspace()
        {
            try
            {
                IWorkspaceName pWorkspaceName = new WorkspaceNameClass();
                pWorkspaceName.WorkspaceFactoryProgID = 
                    "esriDataSourcesGDB.AccessWorkspaceFactory";
                pWorkspaceName.PathName = _personalDatabsePath;
                var n = pWorkspaceName as ESRI.ArcGIS.esriSystem.IName;
                var workspace = n.Open() as IFeatureWorkspace;
                return workspace;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public ESRI.ArcGIS.Geodatabase.IFeatureClass OpenFeatureClasss(string featureClassName)
        {
            var pFeatureWorkspace = OpenWorkspace();
            try
            {
                return pFeatureWorkspace == null ? 
                    null : pFeatureWorkspace.OpenFeatureClass(featureClassName);                
            }
            catch (Exception)
            {
                return null;
            }
        }

        public int FindField(string featureClassName, string fieldName)
        {
            var pFeatrueClass = OpenFeatureClasss(featureClassName);
            return pFeatrueClass == null ? -1 : pFeatrueClass.FindField(fieldName);
        }

        public bool AddField(string featureClassName, string fieldName, int fieldLength, ESRI.ArcGIS.Geodatabase.esriFieldType fieldType)
        {
            var pFeatureClass = OpenFeatureClasss(featureClassName);
            bool flag;
            try
            {
                if (pFeatureClass == null)
                {
                    flag = false;
                }
                else
                {
                    //如果该要素类的字段已经存在
                    if (pFeatureClass.Fields.FindField(fieldName) != -1)
                    {
                        flag = false;
                    }
                    else
                    {
                        var pField = new FieldClass();
                        var pFieldEdit = pField as IFieldEdit;
                        pFieldEdit.Name_2 = fieldName;
                        pFieldEdit.Type_2 = fieldType;
                        pFieldEdit.Length_2 = fieldLength;
                        pFeatureClass.AddField(pFieldEdit);
                        flag = true;
                    }
                    ReleaseFeautureClass(pFeatureClass);
                }
            }
            catch (Exception)
            {
                flag = false;
            }
            return flag;
        }

        public void DeleteIfExist(string feaureClassName)
        {
            IFeatureWorkspace workspace = OpenWorkspace();
            IEnumDataset dataset = (workspace as IWorkspace).get_Datasets(esriDatasetType.esriDTAny);
            IDataset tmp = null;
            while ((tmp = dataset.Next()) != null && tmp.Name != feaureClassName) ;
            if (tmp != null)
                tmp.Delete();
        }

        public void ReleaseFeautureClass(ESRI.ArcGIS.Geodatabase.IFeatureClass pFeatureClass)
        {
            Marshal.FinalReleaseComObject(pFeatureClass);
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
        public bool DeleteFields(ESRI.ArcGIS.Geodatabase.IFeatureClass pFeatureClass, List<ESRI.ArcGIS.Geodatabase.IField> pFields)
        {
            try
            {
                foreach (var deleteField in pFields)
                {
                    pFeatureClass.DeleteField(deleteField);
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
