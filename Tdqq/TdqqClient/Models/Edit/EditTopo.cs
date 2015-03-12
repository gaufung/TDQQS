using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Geometry;
using TdqqClient.Services.AE;
using TdqqClient.Services.Check;
using TdqqClient.Views;

namespace TdqqClient.Models.Edit
{
    class EditTopo:EditModel
    {
        public EditTopo(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }

        public override void Edit(object parameter)
        {
            var ret = true;
            ret &= CreateTopology();
            ret &= TopoCheck();
            if (ret)
            {
                MessageBox.Show(null, "拓扑检查成功", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(null, "拓扑检查错误", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool CreateTopology()
        {
            //获取拓扑名称
            IAeFactory pAeFactory = new PersonalGeoDatabase(PersonDatabase);
            var featureclassNames = pAeFactory.OpenWorkspace().FeatrueClassNames();
            string topoName = string.Empty;
            foreach (var featureclassName in featureclassNames)
            {
                if (featureclassName.EndsWith("Topology"))
                {
                    topoName = featureclassName;
                }
            }
            return string.IsNullOrEmpty(topoName)
                 ? ValidateTopoNotExist(PersonDatabase, SelectFeature)
                 : ValidateTopoExist(PersonDatabase, SelectFeature, topoName);
        }
        private bool ValidateTopoNotExist(string personDatabase, string selectFeaure)
        {
            IAeFactory pAeFactory = new PersonalGeoDatabase(personDatabase);
            IFeatureClass pFeatureClass = pAeFactory.OpenFeatureClasss(selectFeaure);
            IFeatureDataset pDataset = pFeatureClass.FeatureDataset;
            var schemaLock = (ISchemaLock)pDataset;
            bool flag;
            try
            {
                ITopologyContainer2 topologyContainer = (ITopologyContainer2)pDataset;
                schemaLock.ChangeSchemaLock(esriSchemaLock.esriExclusiveSchemaLock);
                ITopology topology = topologyContainer.CreateTopology(selectFeaure + "_Topology",
                    topologyContainer.DefaultClusterTolerance, -1, "");
                topology.AddClass(pFeatureClass, 5, 1, 1, false);
                ITopologyRule topologyRule = new TopologyRuleClass();
                topologyRule.TopologyRuleType = esriTopologyRuleType.esriTRTAreaNoOverlap;
                topologyRule.Name = "Over Lapped";
                topologyRule.OriginClassID = pFeatureClass.FeatureClassID;
                topologyRule.AllOriginSubtypes = true;
                var topologyRuleContainer = (ITopologyRuleContainer)topology;
                topologyRuleContainer.AddRule(topologyRule);
                IEnvelope pEnvelope = GetTopoEnvelope(personDatabase, selectFeaure);
                topology.ValidateTopology(pEnvelope);
                flag = true;
            }
            catch (Exception)
            {
                flag = false;
            }
            finally
            {
                schemaLock.ChangeSchemaLock(esriSchemaLock.esriSharedSchemaLock);
            }
            return flag;
        }
        IEnvelope GetTopoEnvelope(string personDatabase, string selectFeaure)
        {
            IAeFactory pAeFactory = new PersonalGeoDatabase(personDatabase);
            IFeatureClass pFeatureClass = pAeFactory.OpenFeatureClasss(selectFeaure);
            IFeatureCursor pFeatureCursor = pFeatureClass.Search(null, false);
            IFeature pFeature = pFeatureCursor.NextFeature();
            IEnvelope pEnvelope = pFeature.Shape.Envelope; ;
            while (pFeature != null)
            {
                pEnvelope.Union(pFeature.Shape.Envelope);
                pFeature = pFeatureCursor.NextFeature();
            }
            Marshal.ReleaseComObject(pFeatureCursor);
            return pEnvelope;
        }
        private bool ValidateTopoExist(string personDatabase, string selectFeautre, string topoName)
        {

            IAeFactory pAeFactory = new PersonalGeoDatabase(personDatabase);
            IFeatureClass pFeatureClass = pAeFactory.OpenFeatureClasss(selectFeautre);
            IFeatureDataset pDataset = pFeatureClass.FeatureDataset;
            ISchemaLock schemaLock = (ISchemaLock)pDataset;
            bool flag;
            try
            {
                ITopologyContainer2 topologyContainer = (ITopologyContainer2)pDataset;
                schemaLock.ChangeSchemaLock(esriSchemaLock.esriExclusiveSchemaLock);
                var topology = topologyContainer.get_TopologyByName(topoName);
                IEnvelope pEnvelope = GetTopoEnvelope(personDatabase, selectFeautre);
                topology.ValidateTopology(pEnvelope);
                flag = true;
            }
            catch (Exception)
            {
                flag = false;
            }
            finally
            {
                schemaLock.ChangeSchemaLock(esriSchemaLock.esriSharedSchemaLock);
            }
            return flag;
        }
        private bool TopoCheck()
        {
            Hashtable para = new Hashtable();
            Wait wait = new Wait();
            wait.SetWaitCaption("检查拓扑重叠");
            para["w"] = wait;
            para["ret"] = false;
            Thread t = new Thread(new ParameterizedThreadStart(TopoCheck));
            t.Start(para);
            wait.ShowDialog();
            t.Abort();
            return (bool)para["ret"];

        }
        /// <summary>
        /// 线程调用
        /// </summary>
        /// <param name="p"></param>
        private void TopoCheck(object p)
        {
            Hashtable para = p as Hashtable;
            var wait = para["w"] as Wait;
            IAeFactory pAeFactory = new PersonalGeoDatabase(PersonDatabase);
            IFeatureClass pFeatureClass = pAeFactory.OpenFeatureClasss(SelectFeature);
            var count = pFeatureClass.Count();
            try
            {
                IFeatureCursor pFeatureCursor = pFeatureClass.Search(null, false);
                IFeature pFeature;
                int currentIndex = 0;
                while ((pFeature = pFeatureCursor.NextFeature()) != null)
                {
                    wait.SetProgress(((double)currentIndex++ / (double)count));
                    var topoGeometry = pFeature.Shape;
                    ISpatialFilter pSpatialFilter = new SpatialFilterClass();
                    pSpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelOverlaps;
                    pSpatialFilter.Geometry = topoGeometry;
                    IFeatureCursor mFeatureCursor = pFeatureClass.Search(pSpatialFilter, false);
                    IFeature feature = mFeatureCursor.NextFeature();
                    if (feature != null)
                    {
                        para["ret"] = false;
                        wait.CloseWait();
                        return;
                    }
                    Marshal.ReleaseComObject(mFeatureCursor);
                }
                para["ret"] = true;
            }
            catch (Exception)
            {
                para["ret"] = true;
            }
            finally
            {
                pAeFactory.ReleaseFeautureClass(pFeatureClass);
                wait.CloseWait();
            }

        }   
    }
}
