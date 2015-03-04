using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ESRI.ArcGIS.DataSourcesGDB;
using ESRI.ArcGIS.Geodatabase;

namespace TdqqClient.Services.AE
{
    class AeHelper
    {
        /// <summary>
        /// 获取整个个人地理数据库的要素类名称
        /// </summary>
        /// <param name="personDatabase"></param>
        /// <returns></returns>
        public static IEnumerable<string> GetAllFeautureClass(string personDatabase)
        {
            IWorkspaceFactory pWsFt = new AccessWorkspaceFactoryClass();
            IWorkspace pWs = pWsFt.OpenFromFile(personDatabase, 0);
            IEnumDataset pEDataset = pWs.get_Datasets(esriDatasetType.esriDTAny);
            IDataset pDataset;
            while ((pDataset= pEDataset.Next())!= null)
            {
                if (pDataset.Type == esriDatasetType.esriDTFeatureClass) yield return pDataset.Name;
                else if (pDataset.Type == esriDatasetType.esriDTFeatureDataset)
                {
                    IEnumDataset pESubDataset = pDataset.Subsets;
                    IDataset pSubDataset;
                    while ((pSubDataset = pESubDataset.Next()) != null) yield return pSubDataset.Name;
                }
            }
        }

        /// <summary>
        /// 获取某个字段的序号值(小写)
        /// </summary>
        /// <param name="pFeatureClass">要素类</param>
        /// <param name="fieldName">字段名称</param>
        /// <returns>字段的序号值</returns>
        public static int FindFieldIndexInLowerCapper(IFeatureClass pFeatureClass, string fieldName)
        {
            int index = -1;
            for (int i = 0; i < pFeatureClass.Fields.FieldCount; i++)
            {
                if (pFeatureClass.Fields.Field[i].Name.ToLower() == fieldName.ToLower())
                {
                    index = i;
                    break;
                }
            }
            return index;
        }

    }
}
