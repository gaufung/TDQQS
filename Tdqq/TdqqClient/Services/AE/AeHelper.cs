using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ESRI.ArcGIS.DataSourcesGDB;
using ESRI.ArcGIS.Display;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Geometry;
using TdqqClient.Models;
using TdqqClient.Services.Database;

namespace TdqqClient.Services.AE
{
    static  class AeHelper
    {
        /// <summary>
        /// 获取整个个人地理数据库的要素类名称
        /// </summary>
        /// <param name="personDatabase"></param>
        /// <returns></returns>
        //public static List<string> GetAllFeautureClass(string personDatabase)
        //{
        //    IWorkspaceFactory pWsFt = new AccessWorkspaceFactoryClass();
        //    IWorkspace pWs = pWsFt.OpenFromFile(personDatabase, 0);
        //    IEnumDataset pEDataset = pWs.get_Datasets(esriDatasetType.esriDTAny);
        //    List<string> featureList=new List<string>();
        //    IDataset pDataset;
        //    while ((pDataset= pEDataset.Next())!= null)
        //    {
        //        if (pDataset.Type == esriDatasetType.esriDTFeatureClass) featureList.Add(pDataset.Name); 
        //        else if (pDataset.Type == esriDatasetType.esriDTFeatureDataset)
        //        {
        //            IEnumDataset pESubDataset = pDataset.Subsets;
        //            IDataset pSubDataset;
        //            while ((pSubDataset = pESubDataset.Next()) != null) featureList.Add(pSubDataset.Name);
        //        }
        //    }
        //    return featureList;
        //}

        /// <summary>
        /// 扩展方法，获取某个要素空间的所有要素类的名称
        /// </summary>
        /// <param name="pWorkspace"></param>
        /// <returns>要素类名称的集合</returns>
        public static List<string> FeatrueClassNames(this IWorkspace pWorkspace)
        {
            IEnumDataset pEDataset = pWorkspace.get_Datasets(esriDatasetType.esriDTAny);
            List<string> featureList = new List<string>();
            IDataset pDataset;
            while ((pDataset = pEDataset.Next()) != null)
            {
                if (pDataset.Type == esriDatasetType.esriDTFeatureClass) featureList.Add(pDataset.Name);
                else if (pDataset.Type == esriDatasetType.esriDTFeatureDataset)
                {
                    IEnumDataset pESubDataset = pDataset.Subsets;
                    IDataset pSubDataset;
                    while ((pSubDataset = pESubDataset.Next()) != null) featureList.Add(pSubDataset.Name);
                }
            }
            return featureList;
        } 

        /// <summary>
        /// 扩展方法
        /// </summary>
        /// <param name="pFeatureClass">要素类</param>
        /// <param name="fieldName">字段名称</param>
        /// <returns>字段的序号值</returns>
        public static int FindFieldIndexInLowerCapper(this IFeatureClass pFeatureClass, string fieldName)
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

        public static  IEnumerable<SortEnity<object>> GetDks(this IFeatureClass pFeatureClass)
        {
            //获取ObjectID的序号值
            var objectIdIndex = pFeatureClass.Fields.FindField("OBJECTID");
            IFeatureCursor pCursor = pFeatureClass.Search(null, false);
            List<SortEnity<object>> dks = new List<SortEnity<object>>();
            IFeature pFeature = null;
            try
            {
                //循环，获取整体的地块
                while ((pFeature = pCursor.NextFeature()) != null)
                {
                    IEnvelope pEnvelop = pFeature.Shape.Envelope;
                    dks.Add(new SortEnity<object>()
                    {
                        Id = (object)pFeature.get_Value(objectIdIndex),
                        Xcor = pEnvelop.XMax * 0.5 + pEnvelop.XMin * 0.5,
                        Ycor = pEnvelop.YMax * 0.5 + pEnvelop.YMin * 0.5
                    });
                }
            }
            catch (Exception)
            {
                return null;
            }
            return dks;
        }

        public static IRgbColor GetRgb(int r, int g, int b)
        {
            IRgbColor pRgbColor = new RgbColorClass();
            pRgbColor.Red = r;
            pRgbColor.Green = g;
            pRgbColor.Blue = b;
            return pRgbColor;
        }
    }
}
