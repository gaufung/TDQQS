using System;
using System.Collections;
using System.Collections.Generic;
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
    class EditSz:EditModel
    {
        public EditSz(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }

        public override void Edit(object parameter)
        {
            if (!CheckEditFieldsExist())
            {
                MessageBox.Show(null, "字段尚未添加成功", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            IAeFactory pAeFactory = new PersonalGeoDatabase(PersonDatabase);
            if (!pAeFactory.IsExist(SelectFeature + "_JZD") || !pAeFactory.IsExist(SelectFeature + "_JZX"))
            {
                MessageBox.Show(null, "尚未提取界址线或界址点", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (CreateSz())
            {
                MessageBox.Show(null, "四至提取成功", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(null, "四至提取失败", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool CreateSz()
        {
            Wait wait = new Wait();
            wait.SetWaitCaption("正在提取四至");
            Hashtable para = new Hashtable()
            {
                {"wait",wait},
                {"ret",false}
            };
            Thread t = new Thread(new ParameterizedThreadStart(CreateSz));
            t.Start(para);
            wait.ShowDialog();
            t.Abort();
            return (bool)para["ret"];
        }

        private void CreateSz(object p)
        {
            Hashtable para = p as Hashtable;
            Wait wait = para["wait"] as Wait;
            IAeFactory pAeFactory = new PersonalGeoDatabase(PersonDatabase);
            IFeatureClass inputFC = pAeFactory.OpenFeatureClasss(SelectFeature);
            int total = inputFC.Count();
            var pDataset = inputFC as IDataset;
            IWorkspaceEdit workspaceEdit = pDataset.Workspace as IWorkspaceEdit;
            var dxnbIndex = GetDxnbIndex(inputFC);
            workspaceEdit.StartEditing(false);
            workspaceEdit.StartEditOperation();
            IFeatureCursor featureCursor = inputFC.Update(null, false);
            IFeature feature = featureCursor.NextFeature();
            int currentIndex = 0;
            List<string> idList=new List<string>();
            while (feature != null)
            {
                //存储四至的数组
                wait.SetProgress(((double)currentIndex++ / (double)total));
                string[] szArr = { "路", "路", "路", "路" };
                szArr = GetOneFeaSZ(feature, PersonDatabase, SelectFeature + "_JZX");
                var objectid = feature.Value[feature.Fields.FindField("OBJECTID")].ToString();
                if (szArr == null)
                {
                    /*
                    MessageBox.Show(null, "请做一下拓扑检查",
                        "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    wait.CloseWait();
                    para["ret"] = false;
                    return;
                     */
                    idList.Add(objectid);
                    feature = featureCursor.NextFeature();
                    continue;
                }
                if (!feature.get_Value(dxnbIndex[0]).ToString().StartsWith("_"))
                {
                    feature.set_Value(dxnbIndex[0], szArr[0]);
                }
                if (!feature.get_Value(dxnbIndex[1]).ToString().StartsWith("_"))
                {
                    feature.set_Value(dxnbIndex[1], szArr[1]);
                }
                if (!feature.get_Value(dxnbIndex[2]).ToString().StartsWith("_"))
                {
                    feature.set_Value(dxnbIndex[2], szArr[2]);
                }
                if (!feature.get_Value(dxnbIndex[3]).ToString().StartsWith("_"))
                {
                    feature.set_Value(dxnbIndex[3], szArr[3]);
                }
                featureCursor.UpdateFeature(feature);
                feature = featureCursor.NextFeature();
            }
            featureCursor.Flush();
            Marshal.ReleaseComObject(featureCursor);
            workspaceEdit.StopEditOperation();
            workspaceEdit.StopEditing(true);
            MessageBox.Show(idList.Count.ToString());
            wait.CloseWait();
            para["ret"] = true;
        }
        private List<int> GetDxnbIndex(IFeatureClass pFeatureClass)
        {
            List<int> dxnzList = new List<int>();
            int index;
            index = pFeatureClass.Fields.FindField("DKDZ");
            if (index != -1) dxnzList.Add(index);
            index = pFeatureClass.Fields.FindField("DKNZ");
            if (index != -1) dxnzList.Add(index);
            index = pFeatureClass.Fields.FindField("DKXZ");
            if (index != -1) dxnzList.Add(index);
            index = pFeatureClass.Fields.FindField("DKBZ");
            if (index != -1) dxnzList.Add(index);
            return dxnzList;
        }
        private string[] GetOneFeaSZ(IFeature feature, string databaseUrl, string jzxFeature)
        {
            IPolygon polygon = feature.ShapeCopy as IPolygon;
            //拿出所有与地块相交的界址线
            IAeFactory pAeFactory = new PersonalGeoDatabase(databaseUrl);
            IFeatureClass jzxFeatureClass = pAeFactory.OpenFeatureClasss(jzxFeature);
            SpatialFilter sf = new SpatialFilterClass();
            sf.GeometryField = "SHAPE";
            sf.Geometry = polygon;
            sf.SpatialRel = esriSpatialRelEnum.esriSpatialRelRelation;
            sf.SpatialRelDescription = "FFTTT*FF*";
            IFeatureCursor jzxCursor = jzxFeatureClass.Search(sf, false);
            List<IFeature> jzxList = new List<IFeature>();
            //存储东西走向与南北走向的边的序号
            IFeature tmpFeature = null;
            while ((tmpFeature = jzxCursor.NextFeature()) != null)
            {
                jzxList.Add(tmpFeature);
            }
            Marshal.ReleaseComObject(jzxCursor);
            //拿出所有的定点
            IPointCollection pointCollection = polygon as IPointCollection;
            //并拿出最东南西北的点(x最小值最大值，Y最小值做大值)
            int d = 0, n = 0, x = 0, b = 0;
            double xMax = pointCollection.get_Point(0).X;
            double yMin = pointCollection.get_Point(0).Y;
            double xMin = pointCollection.get_Point(0).X;
            double yMax = pointCollection.get_Point(0).Y;
            for (int i = 1; i < pointCollection.PointCount - 1; i++)
            {
                if (pointCollection.get_Point(i).X > xMax)
                {
                    xMax = pointCollection.get_Point(i).X;
                    d = i;
                }
                if (pointCollection.get_Point(i).Y < yMin)
                {
                    yMin = pointCollection.get_Point(i).Y;
                    n = i;
                }
                if (pointCollection.get_Point(i).X < xMin)
                {
                    xMin = pointCollection.get_Point(i).X;
                    x = i;
                }
                if (pointCollection.get_Point(i).Y > yMax)
                {
                    yMax = pointCollection.get_Point(i).Y;
                    b = i;
                }
            }
            //拿出对应的四个点
            IPoint dPoint = pointCollection.get_Point(d);
            IPoint nPoint = pointCollection.get_Point(n);
            IPoint xPoint = pointCollection.get_Point(x);
            IPoint bPoint = pointCollection.get_Point(b);
            //拿出对应的边放入List
            List<IFeature> dList = new List<IFeature>();
            List<IFeature> nList = new List<IFeature>();
            List<IFeature> xList = new List<IFeature>();
            List<IFeature> bList = new List<IFeature>();
            for (int i = 0; i < jzxList.Count; i++)
            {
                //IPolyline jzdPolyline = jzxList[i].ShapeCopy as IPolyline;
                ISegmentCollection sc = jzxList[i].ShapeCopy as ISegmentCollection;
                ISegment jzxSegment = sc.get_Segment(0);
                int cc = 0;
                double A = jzxSegment.ToPoint.X;
                if (IsTwoPointEqual(jzxSegment.ToPoint, dPoint) || IsTwoPointEqual(jzxSegment.FromPoint, dPoint))
                {
                    dList.Add(jzxList[i]);
                }
                if (IsTwoPointEqual(jzxSegment.ToPoint, nPoint) || IsTwoPointEqual(jzxSegment.FromPoint, nPoint))
                {
                    nList.Add(jzxList[i]);
                }
                if (IsTwoPointEqual(jzxSegment.ToPoint, xPoint) || IsTwoPointEqual(jzxSegment.FromPoint, xPoint))
                {
                    xList.Add(jzxList[i]);
                }
                if (IsTwoPointEqual(jzxSegment.ToPoint, bPoint) || IsTwoPointEqual(jzxSegment.FromPoint, bPoint))
                {
                    bList.Add(jzxList[i]);
                }
            }
            //判断是不是每个点都有两个边与其相连
            if ((dList.Count >= 2 && nList.Count >= 2 && xList.Count >= 2 && bList.Count >= 2) == false)
            {
                return null;
            }
            //判断list中那条边是真正的东南西北至
            //判断东至
            IFeature dFeature = null;
            if (IsH(dList[0]) && IsS(dList[1]))
            {
                dFeature = dList[1];
            }
            else if (IsH(dList[1]) && IsS(dList[0]))
            {
                dFeature = dList[0];
            }
            else
            {
                if (CalSlope(dList[0]) >= CalSlope(dList[1]))
                {
                    dFeature = dList[0];
                }
                else
                {
                    dFeature = dList[1];
                }
            }
            //判断南至
            IFeature nFeature = null;
            if (IsH(nList[0]) && IsS(nList[1]))
            {
                nFeature = nList[0];
            }
            else if (IsH(nList[1]) && IsS(nList[0]))
            {
                nFeature = nList[1];
            }
            else
            {
                if (CalSlope(nList[0]) <= CalSlope(nList[1]))
                {
                    nFeature = nList[0];
                }
                else
                {
                    nFeature = nList[1];
                }
            }
            //判断xi至
            IFeature xFeature = null;
            if (IsH(xList[0]) && IsS(xList[1]))
            {
                xFeature = xList[1];
            }
            else if (IsH(xList[1]) && IsS(xList[0]))
            {
                xFeature = xList[0];
            }
            else
            {
                if (CalSlope(xList[0]) >= CalSlope(xList[1]))
                {
                    xFeature = xList[0];
                }
                else
                {
                    xFeature = xList[1];
                }
            }
            //判断北至
            IFeature bFeature = null;
            if (IsH(bList[0]) && IsS(bList[1]))
            {
                bFeature = bList[0];
            }
            else if (IsH(bList[1]) && IsS(bList[0]))
            {
                bFeature = bList[1];
            }
            else
            {
                if (CalSlope(bList[0]) <= CalSlope(bList[1]))
                {
                    bFeature = bList[0];
                }
                else
                {
                    bFeature = bList[1];
                }
            }
            //最终确定东西南北至
            string[] szArr = new string[4];
            string CBFMC = feature.get_Value(feature.Fields.FindField("CBFMC")).ToString();
            string PLDWQLR = "";
            string[] QLRARR = null;
            //dongxinanbei
            PLDWQLR = dFeature.get_Value(dFeature.Fields.FindField("PLDWQLR")).ToString();
            QLRARR = PLDWQLR.Split(',');
            if (QLRARR[0].Equals(CBFMC))
            {
                szArr[0] = QLRARR[1];
            }
            else
            {
                szArr[0] = QLRARR[0];
            }
            //NAN
            PLDWQLR = nFeature.get_Value(dFeature.Fields.FindField("PLDWQLR")).ToString();
            QLRARR = PLDWQLR.Split(',');
            if (QLRARR[0].Equals(CBFMC))
            {
                szArr[1] = QLRARR[1];
            }
            else
            {
                szArr[1] = QLRARR[0];
            }
            //xi
            PLDWQLR = xFeature.get_Value(dFeature.Fields.FindField("PLDWQLR")).ToString();
            QLRARR = PLDWQLR.Split(',');
            if (QLRARR[0].Equals(CBFMC))
            {
                szArr[2] = QLRARR[1];
            }
            else
            {
                szArr[2] = QLRARR[0];
            }
            //bei
            PLDWQLR = bFeature.get_Value(dFeature.Fields.FindField("PLDWQLR")).ToString();
            QLRARR = PLDWQLR.Split(',');
            if (QLRARR[0].Equals(CBFMC))
            {
                szArr[3] = QLRARR[1];
            }
            else
            {
                szArr[3] = QLRARR[0];
            }
            //遍历四至数组，如果“”就赋值为路
            for (int i = 0; i < 4; i++)
            {
                if (szArr[i].Equals(""))
                    szArr[i] = "路";
            }
            return szArr;
        }
        private bool IsS(IFeature feature)
        {
            ISegmentCollection sc = feature.ShapeCopy as ISegmentCollection;
            ISegment segment = sc.get_Segment(0);
            double dx = segment.ToPoint.X - segment.FromPoint.X;
            double dy = segment.ToPoint.Y - segment.FromPoint.Y;
            if (dx == 0)
            {
                return true;
            }
            else
            {
                double slope = Math.Abs(dy / dx);
                if (slope > 1)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            return false;
        }
        private bool IsTwoPointEqual(IPoint point1, IPoint point2)
        {
            double tolerance = 0.001d;
            if (Math.Abs(point1.X - point2.X) < tolerance && Math.Abs(point1.Y - point2.Y) < tolerance)
            {
                return true;
            }
            return false;
        }
        private bool IsH(IFeature feature)
        {
            ISegmentCollection sc = feature.ShapeCopy as ISegmentCollection;
            ISegment segment = sc.get_Segment(0);
            double dx = segment.ToPoint.X - segment.FromPoint.X;
            double dy = segment.ToPoint.Y - segment.FromPoint.Y;
            if (dx == 0)
            {
                return false;
            }
            else
            {
                double slope = Math.Abs(dy / dx);
                if (slope <= 1)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            return false;
        }
        private double CalSlope(IFeature feature)
        {
            ISegmentCollection sc = feature.ShapeCopy as ISegmentCollection;
            ISegment segment = sc.get_Segment(0);
            double dx = segment.ToPoint.X - segment.FromPoint.X;
            double dy = segment.ToPoint.Y - segment.FromPoint.Y;
            if (dx == 0)
            {
                return 999999.9;
            }
            else
            {
                return Math.Abs(dy / dx);
            }
        }
    }
}
