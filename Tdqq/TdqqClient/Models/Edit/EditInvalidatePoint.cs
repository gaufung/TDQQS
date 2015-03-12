using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Geometry;
using TdqqClient.Services.AE;
using TdqqClient.Services.Check;
using TdqqClient.Views;

namespace TdqqClient.Models.Edit
{
    class EditInvalidatePoint:EditModel
    {
        public EditInvalidatePoint(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }

        public override void Edit(object parameter)
        {
            if (DeleteInvaildPoint())
            {
                MessageBox.Show(null, "删除无效点成功", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(null, "删除无效点失败", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }   
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
        public bool DeleteInvaildPoint()
        {

            Wait wait = new Wait();
            wait.SetWaitCaption("删除无效的点");
            Hashtable para = new Hashtable();
            para["wait"] = wait;
            para["ret"] = false;
            Thread t = new Thread(new ParameterizedThreadStart(DeleteInvaildPoint));
            t.Start(para);
            wait.ShowDialog();
            t.Abort();
            return (bool)para["ret"];
        }

        private void DeleteInvaildPoint(object p)
        {
            var para = p as Hashtable;
            var wait = para["wait"] as Wait;
            IAeFactory pAeFactory = new PersonalGeoDatabase(PersonDatabase);
            IFeatureClass inputFC = pAeFactory.OpenFeatureClasss(SelectFeature);
            int total = inputFC.Count();
            try
            {
                var pDataset = inputFC as IDataset;
                IWorkspaceEdit workSpaceEdit = pDataset.Workspace as IWorkspaceEdit;
                workSpaceEdit.StartEditing(true);
                workSpaceEdit.StartEditOperation();
                IFeatureCursor featureCursor = inputFC.Update(null, false);
                IFeature feature;
                int currentIndex = 0;
                while ((feature = featureCursor.NextFeature()) != null)
                {
                    //用删除过折点的多边形替换当前多边形
                    wait.SetProgress(((double)currentIndex++ / (double)total));
                    IPolygon polygon = (IPolygon)feature.ShapeCopy;
                    //int pointCount = polygon.
                    IPolygon tempPolygon = DelectOneFsExcessVertex(polygon);
                    feature.Shape = tempPolygon;
                    featureCursor.UpdateFeature(feature);
                }
                featureCursor.Flush();
                Marshal.ReleaseComObject(featureCursor);
                workSpaceEdit.StopEditOperation();
                workSpaceEdit.StopEditing(true);
                para["ret"] = true;
            }
            catch (Exception)
            {
                para["ret"] = false;
            }
            finally
            {
                wait.CloseWait();
                pAeFactory.ReleaseFeautureClass(inputFC);
            }
        }

        //n环Polygon拆分dddddd
        private IPolygon DelectOneFsExcessVertex(IPolygon polygon)
        {
            //判断Polygon有几个环
            IPointCollection pointCollection = polygon as IPointCollection;
            int pCount = pointCollection.PointCount;
            ISegmentCollection segmentCollection = polygon as ISegmentCollection;
            int sCount = segmentCollection.SegmentCount;
            if (pCount == sCount + 1)
            {
                return DelectOneFExcessVertex(polygon);
            }
            else //把polygon拆分
            {
                IGeometryCollection bPolygons = new GeometryBagClass();
                ITopologicalOperator tOperator = new PolygonClass();
                List<IPolygon> polygons = new List<IPolygon>();
                ISegmentCollection tempSegmentCollection = segmentCollection;
                IPolygon tempPolygon = GetPolygonFromSegmentCollection(tempSegmentCollection);
                polygons.Add(DelectOneFExcessVertex(tempPolygon));
                tempSegmentCollection = GetRestSegmentCollectionFromSegmentCollection(segmentCollection);
                while (tempSegmentCollection != null)
                {
                    tempPolygon = GetPolygonFromSegmentCollection(tempSegmentCollection);
                    polygons.Add(DelectOneFExcessVertex(tempPolygon));
                    tempSegmentCollection = GetRestSegmentCollectionFromSegmentCollection(tempSegmentCollection);
                }
                foreach (IPolygon ePolygon in polygons)
                {
                    bPolygons.AddGeometry(ePolygon);
                }
                tOperator.ConstructUnion(bPolygons as IEnumGeometry);
                return tOperator as IPolygon;
            }
        }
        //函数给定一条边和所有边的集合，返回边所在polygonddddddd
        private IPolygon GetPolygonFromSegmentCollection(ISegmentCollection segmentCollection)
        {
            ISegmentCollection resultSegmentCollection = new PolygonClass();
            ISegment tempSegment = segmentCollection.get_Segment(0);
            resultSegmentCollection.AddSegment(tempSegment);
            IPoint startPoint = tempSegment.FromPoint;
            for (int i = 0; i < segmentCollection.SegmentCount; i++)
            {
                //if (tempSegment.ToPoint.X == segmentCollection.get_Segment(i).FromPoint.X && tempSegment.ToPoint.Y == segmentCollection.get_Segment(i).FromPoint.Y)
                if (IsTwoPointEqual(tempSegment.ToPoint, segmentCollection.get_Segment(i).FromPoint))
                {
                    tempSegment = segmentCollection.get_Segment(i);
                    resultSegmentCollection.AddSegment(tempSegment);
                    // if (startPoint.X == tempSegment.ToPoint.X && startPoint.Y == tempSegment.ToPoint.Y)
                    if (IsTwoPointEqual(startPoint, tempSegment.ToPoint))
                    {
                        return resultSegmentCollection as IPolygon;
                    }
                }
            }
            return null;
        }
        private ISegmentCollection GetRestSegmentCollectionFromSegmentCollection(ISegmentCollection segmentCollection)
        {
            List<int> savedIndex = new List<int>();
            ISegmentCollection resultSegmentCollection = new PolygonClass();
            ISegment tempSegment = segmentCollection.get_Segment(0);
            // resultSegmentCollection.AddSegment(tempSegment);
            savedIndex.Add(0);
            IPoint startPoint = tempSegment.FromPoint;
            for (int i = 0; i < segmentCollection.SegmentCount; i++)
            {
                //  if (tempSegment.ToPoint.X == segmentCollection.get_Segment(i).FromPoint.X && tempSegment.ToPoint.Y == segmentCollection.get_Segment(i).FromPoint.Y)
                if (IsTwoPointEqual(tempSegment.ToPoint, segmentCollection.get_Segment(i).FromPoint))
                {
                    tempSegment = segmentCollection.get_Segment(i);
                    //resultSegmentCollection.AddSegment(tempSegment);
                    savedIndex.Add(i);
                    // if (startPoint.X == tempSegment.ToPoint.X && startPoint.Y == tempSegment.ToPoint.Y)
                    if (IsTwoPointEqual(startPoint, tempSegment.ToPoint))
                    {
                        //选择没有用过的Segment,r如果全部已经存储过，返回Null
                        if (savedIndex.Count == segmentCollection.SegmentCount)
                        {
                            return null;
                        }
                        else
                        {
                            for (int j = 0; j < segmentCollection.SegmentCount; j++)
                            {
                                if (IsIndexInList(j, savedIndex) == false)
                                {
                                    resultSegmentCollection.AddSegment(segmentCollection.get_Segment(j));
                                }
                            }
                            return resultSegmentCollection;
                        }

                    }
                }
            }
            return null;
        }
        //Polygon（单环）返回消除过折点的Polygonddddd
        private IPolygon DelectOneFExcessVertex(IPolygon polygon)
        {
            //判断Polygon有几个环
            IPointCollection pointCollection = polygon as IPointCollection;
            int pCount = pointCollection.PointCount;
            //记录需要删除的点的index
            List<int> listDelectVertexIndex = new List<int>();
            double[] slopeArr = new double[pCount - 1];
            for (int i = 0; i < pCount - 1; i++)
            {
                IPoint startPoint = pointCollection.get_Point(i);
                IPoint endPoint = pointCollection.get_Point(i + 1);
                double sX = startPoint.X;
                double sY = startPoint.Y;
                double eX = endPoint.X;
                double eY = endPoint.Y;
                double dx = eX - sX;
                double dy = eY - sY;
                // if (TPT(eX) - TPT(sX) == 0.0)
                if (Math.Abs(TPT(eX) - TPT(sX)) < 0.0001d)
                {
                    slopeArr[i] = 1000000000;
                }
                else
                {
                    slopeArr[i] = dy / dx;
                }
            }
            //判断如果前后两条边斜率相等，删除中间点
            for (int i = 0; i < pCount - 2; i++)
            {
                //if (slopeArr[i + 1] == 0.0)
                if (Math.Abs(slopeArr[i + 1]) < 0.0001d)
                {
                    //if (slopeArr[i] == 0.0)
                    if (Math.Abs(slopeArr[i]) < 0.0001d)
                    {
                        listDelectVertexIndex.Add(i + 1);
                    }
                }
                else
                {
                    if (SlopeIsSame(slopeArr[i], slopeArr[i + 1]))
                    {
                        listDelectVertexIndex.Add(i + 1);
                    }
                }

            }
            //if (slopeArr[0] == 0.0)
            if (Math.Abs(slopeArr[0]) < 0.0001d)
            {
                //if (slopeArr[pCount - 2] == 0.0)
                if (Math.Abs(slopeArr[pCount - 2]) < 0.0001d)
                {
                    listDelectVertexIndex.Add(0);
                }
            }
            else
            {
                if (SlopeIsSame(slopeArr[pCount - 2], slopeArr[0]))
                {
                    listDelectVertexIndex.Add(0);
                }
            }
            /**/
            IPointCollection finalPointCollection = new PolygonClass();
            //构建新的点集
            for (int i = 0; i < pointCollection.PointCount - 1; i++)
            {
                if (IsIndexInList(i, listDelectVertexIndex) == false)
                {
                    finalPointCollection.AddPoint(pointCollection.get_Point(i));
                }
            }
            //闭合
            IPoint tempPoint = finalPointCollection.get_Point(0);
            finalPointCollection.AddPoint(tempPoint);
            IPolygon resultPolygon = finalPointCollection as IPolygon;
            return resultPolygon;
        }
        //判断录入的两个斜率是否相等，不包含分母等于零的情况
        private Boolean SlopeIsSame(double sSlope, double eSlope)
        {
            //不同正负号，直接返回False
            if (sSlope / eSlope < 0)
            {
                return false;
            }
            else if (Math.Abs(sSlope) < 1 && Math.Abs(eSlope) < 1)//----
            {
                if (Math.Abs(sSlope / eSlope - 1) < 0.01)//同号相除比为正
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else///|||///
            {
                if (Math.Abs(sSlope / eSlope - 1) < 0.001)
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
        //ddddddddddd
        private double TPT(double num)
        {
            string tempStr = num.ToString() + "0000";
            int index = tempStr.IndexOf('.');
            double TPT = Convert.ToDouble(tempStr.Substring(0, index + 4));
            return TPT;
        }
        //ddddddddd
        private Boolean IsIndexInList(int index, List<int> list)
        {
            var items = from item in list
                        where item == index
                        select item;
            return items.Any();
        }   
    }
}
