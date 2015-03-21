using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using ESRI.ArcGIS.ADF.COMSupport;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Display;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Output;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using TdqqClient.Services.AE;
using TdqqClient.Services.Database;
using TdqqClient.Services.Export;

namespace TdqqClient.Models.Export.ExportSingle
{
    internal class MapExport : ExportBase
    {
        public MapExport(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {
        }

        public override void Export(string cbfmc, string cbfbm, string folderPath, string edgeFeature = "")
        {
            var saveFilePath = folderPath + @"\" + cbfbm.Substring(14) + "_" + cbfmc + "_04地块示意图.xls";
            var fbfdz = SelectFbfInfo()[6].ToString().Trim();
            ExportDksyt(saveFilePath,cbfbm,edgeFeature,cbfmc,fbfdz);
            Export2Pdf.Excel2Pdf(saveFilePath);
            File.Delete(saveFilePath);

        }



        private void ExportDksyt(string toSaveFilePath, string cbfbm, string cunEdge, string cbfmc, string fbfdz)
        {
           
                var sqlString = string.Format("Select SCMJ from {0} where trim(CBFBM)='{1}'", SelectFeature, cbfbm);
                var accessFactory = new MsAccessDatabase(PersonDatabase);
                var dtField = accessFactory.Query(sqlString);
                double scmj = 0.0;
                for (int k = 0; k < dtField.Rows.Count; k++)
                {
                    scmj += Convert.ToDouble(double.Parse(dtField.Rows[k][0].ToString().Trim()).ToString("f"));
                }
                sqlString = string.Format("select CBFCYSL from {0} where trim(CBFBM)='{1}'", "CBF", cbfbm);
                accessFactory = new MsAccessDatabase(BasicDatabase);
                var dt = accessFactory.Query(sqlString);
                IAeFactory aeFactory = new PersonalGeoDatabase(PersonDatabase);
                IFeatureClass zdtFC = aeFactory.OpenFeatureClasss(SelectFeature);
                IFeatureClass bjxFC = aeFactory.OpenFeatureClasss(cunEdge);
                IMapDocument mapDoc = new MapDocumentClass();
                mapDoc.Open(AppDomain.CurrentDomain.BaseDirectory + @"\dkct.mxd", "");
                IMap pMap = mapDoc.get_Map(0);
                var templatePath = GetTemplatePath(cbfbm);
                File.Copy(templatePath, toSaveFilePath, true);
                ExportOneSyt(pMap, cbfbm, zdtFC, bjxFC, toSaveFilePath, PersonDatabase, cbfmc,
                    dt.Rows[0][0].ToString(), dtField.Rows.Count, scmj, fbfdz);
           
        }

        private string GetTemplatePath(string cbfbm)
        {
            var dt = SelectFieldsByCbfbm(cbfbm);
            const int perSheetImage = 4;
            return AppDomain.CurrentDomain.BaseDirectory + @"\template\地块分布图\梁山土地承包地块示意图" +
                   Math.Ceiling((decimal) dt.Rows.Count/(decimal) perSheetImage).ToString() + @".xls";
        }

        private  void ExportOneSyt(IMap map, string cbfbm, IFeatureClass zdtFC, IFeatureClass bjxFC, string xlsPath,
            string dbUrl,
            string cbfmc, string familyNumber, int fieldNumber, double scmj, string address)
        {
            if (map == null || cbfbm == null || zdtFC == null || bjxFC == null) return;
            IFeatureClass new_zdtFC = null;
            List<DkInfo> dkInfoList = null;
            IFeatureClass pointFC = null;
            if (!BuildNewFeatureClass(zdtFC, cbfbm, out new_zdtFC, out pointFC, out dkInfoList, dbUrl))
            {
               throw new Exception("创建地块要素失败");
            }
            if (!InsertPointClass(new_zdtFC, dkInfoList, pointFC))
            {
                throw new Exception("插入点要素失败");
            }
            

            if (!FixDataSource(map, new_zdtFC, bjxFC, pointFC))return;


            if (!SetQuery(map, cbfbm)) return;

            if (!AddRoadTextElement(map, new_zdtFC, dkInfoList, pointFC)) return;
            IEnvelope envelope = null;
            IFeatureCursor bjxCursor = bjxFC.Search(null, false);
            IFeature tmpF = bjxCursor.NextFeature();
            ILayer zdtLayer = null;
            ILayer bjxLayer = null;
            ILayer pointLayer = null;
            ILayer layer = null;
            for (int i = 0; i < map.LayerCount; ++i)
            {
                layer = map.get_Layer(i);
                if (layer.Name == "zdt")
                {
                    zdtLayer = layer;
                }
                else if (layer.Name == "bjx")
                {
                    bjxLayer = layer;
                }
                else if (layer.Name == "point")
                {
                    pointLayer = layer;
                }
            }
            if (zdtLayer == null || bjxLayer == null || pointLayer == null)
            {
                throw new Exception("地图文档内容错误！");
            }



            /////地块详细图
            zdtLayer.Visible = true;
            bjxLayer.Visible = false;
            pointLayer.Visible = false;
            //查询
            IQueryFilter query = new QueryFilterClass();
            query.WhereClause = "CBFBM = \"" + cbfbm + "\"";
            IFeatureCursor cursor = new_zdtFC.Search(query, false);
            IFeature feature;
            int count = 0;
            string syt;
            while ((feature = cursor.NextFeature()) != null)
            {
                //图片位置

                syt = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                    "tmp\\dk" + count.ToString() + ".jpg"); //临时缩略图路径
                if (File.Exists(syt))
                    File.Delete(syt);
                envelope = feature.Shape.Envelope;
                ISpatialFilter pSpatialFilter = new SpatialFilterClass();
                pSpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelTouches;
                pSpatialFilter.Geometry = feature.Shape;
                IFeatureCursor mFeatureCursor = new_zdtFC.Search(pSpatialFilter, false);
                IFeature pFeature = mFeatureCursor.NextFeature();
                while (pFeature != null)
                {
                    envelope.Union(pFeature.Shape.Envelope);
                    pFeature = mFeatureCursor.NextFeature();
                }
                Marshal.ReleaseComObject(mFeatureCursor);
                if (!ExportImageToLocal(map as IActiveView, envelope, 320, 420, syt))
                {
                    throw new Exception("导出图片出错！");
                }
                count++;
            }
            Marshal.ReleaseComObject(cursor);
            IGraphicsContainer gContainer = map as IGraphicsContainer;
            gContainer.DeleteAllElements();

            //缩略图
            zdtLayer.Visible = false;
            bjxLayer.Visible = true;
            pointLayer.Visible = true;
            string slt = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tmp\\slt.jpg"); //临时缩略图路径
            if (File.Exists(slt))
                File.Delete(slt);
            envelope = tmpF.Shape.Envelope;
            if (!ExportImageToLocal(map as IActiveView, envelope, 150, 150, slt))
            {
                throw new Exception("导出图片出错！");
            }

            #region 清除产生临时文件

            for (int i = 0; i < map.LayerCount; i++)
            {
                IGeoFeatureLayer tmp = map.get_Layer(i) as IGeoFeatureLayer;
                tmp.FeatureClass = null;
            }
            gContainer.DeleteAllElements();
            //清空数据
            Marshal.FinalReleaseComObject(new_zdtFC);
            Marshal.FinalReleaseComObject(pointFC);
            new_zdtFC = null;
            pointFC = null;

            GC.WaitForPendingFinalizers();
            GC.Collect();

            #endregion

            const int perCount = 4;
            //插入Excel其他信息和缩略图
            using (var fileStream = new System.IO.FileStream(xlsPath, FileMode.Open, FileAccess.ReadWrite))
            {
                IWorkbook workbook = new HSSFWorkbook(fileStream);
                ISheet sheet = workbook.GetSheetAt(0);
                AddPictureToExcle(slt, workbook, 20, 20, 0, 0, 7, 6, 9, 9, false, 0);
                NPOI.SS.UserModel.IRow row = null;
                row = sheet.GetRow(6);
                row.GetCell(2).SetCellValue(cbfbm.Substring(14, 4));
                row.GetCell(4).SetCellValue(cbfmc);
                row.GetCell(6).SetCellValue(familyNumber + "人");
                row = sheet.GetRow(7);
                row.GetCell(2).SetCellValue(fieldNumber.ToString() + "块");
                row.GetCell(4).SetCellValue(scmj.ToString("f") + "亩");
                row.GetCell(5).SetCellValue("第1张");
                row = sheet.GetRow(8);
                row.GetCell(2).SetCellValue(address);
                row.GetCell(5).SetCellValue("共" + Math.Ceiling((double)count / (double)perCount) + "张");
                System.IO.FileStream fs = new System.IO.FileStream(xlsPath, FileMode.Create, FileAccess.Write);
                workbook.Write(fs);
                fs.Close();
                fileStream.Close();
            }
            //每张sheet文件
            using (var fileStream = new System.IO.FileStream(xlsPath, FileMode.Open, FileAccess.ReadWrite))
            {
                IWorkbook workbook = new HSSFWorkbook(fileStream);
                for (int i = 0; i < count; i++)
                {
                    int sheetIndex = i/perCount;
                    //ISheet sheet = workbook.GetSheetAt(sheetIndex);
                    syt = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                        "tmp\\dk" + i.ToString() + ".jpg");
                    int position = i%perCount;
                    int dx1, dy1, dx2, dy2, col1, row1, col2, row2;
                    dx1 = dy1 = 20;
                    dx2 = dy2 = 10;
                    col1 = 1;
                    row1 = 2;
                    col2 = 5;
                    row2 = 4;
                    if (position == 1)
                    {
                        dx1 = dy1 = 20;
                        dx2 = dy2 = 10;
                        col1 = 5;
                        row1 = 2;
                        col2 = 9;
                        row2 = 4;
                    }
                    if (position == 2)
                    {
                        dx1 = dy1 = 20;
                        dx2 = dy2 = 10;
                        col1 = 1;
                        row1 = 4;
                        col2 = 5;
                        row2 = 6;
                    }
                    if (position == 3)
                    {
                        dx1 = dy1 = 20;
                        dx2 = dy2 = 10;
                        col1 = 5;
                        row1 = 4;
                        col2 = 9;
                        row2 = 6;
                    }
                    AddPictureToExcle(syt, workbook, dx1, dy1, dx2, dy2, col1, row1, col2, row2, true, sheetIndex);
                    if (i%perCount == 0)
                    {
                        AddPostScript(workbook, sheetIndex, cbfbm, cbfmc, familyNumber, fieldNumber, scmj, address,
                            sheetIndex + 1, (int) Math.Ceiling((double) count/(double) perCount));
                    }
                }
                System.IO.FileStream fs = new System.IO.FileStream(xlsPath, FileMode.Create, FileAccess.Write);
                workbook.Write(fs);
                fs.Close();
                fileStream.Close();
            }
        }
        private  void AddPictureToExcle(string picPath, IWorkbook workbook,
           int dx1, int dy1, int dx2, int dy2, int col1, int row1, int col2, int row2, bool isResize, int pageIndex)
        {
            byte[] picByte = File.ReadAllBytes(picPath);
            int picInt = workbook.AddPicture(picByte, PictureType.JPEG);
            ISheet sheet = workbook.GetSheetAt(pageIndex);

            HSSFPatriarch patriarch = (HSSFPatriarch)sheet.CreateDrawingPatriarch();
            HSSFClientAnchor anchor = new HSSFClientAnchor(dx1, dy1, dx2, dy2, col1, row1, col2, row2);
            IPicture pic = patriarch.CreatePicture(anchor, picInt);
            if (isResize) pic.Resize();
            return;
        }
        /// <summary>
        /// 添加路标签到地图上
        /// </summary>
        /// <param name="map">地图文档接口</param>
        /// <param name="new_zdtFC">筛选后新建的宗地图类</param>
        /// <param name="dkInfoList">地块信息列表</param>
        /// <param name="pointFC">宗地图中心点要素类</param>
        /// <returns></returns>
        private  bool AddRoadTextElement(IMap map, IFeatureClass new_zdtFC, List<DkInfo> dkInfoList, IFeatureClass pointFC)
        {
            IGraphicsContainer graphContainer = map as IGraphicsContainer;
            ITextElement road = null;
            road = new TextElementClass();
            IFeature feature = null;
            IPoint point = null;
            IPolygon polygon = new PolygonClass();
            IPointCollection pointCollection = polygon as IPointCollection;

            foreach (DkInfo info in dkInfoList)
            {
                feature = new_zdtFC.GetFeature(info.dkid);

                if (info.dz == DkInfo.ROAD_ID)
                {
                    point = GetTextPoint(feature, 0);

                    road = CreateTextElement(point, esriTextHorizontalAlignment.esriTHALeft, esriTextVerticalAlignment.esriTVACenter);
                    graphContainer.AddElement(road as IElement, 0);
                }

                if (info.xz == DkInfo.ROAD_ID)
                {
                    point = GetTextPoint(feature, 1);
                    road = CreateTextElement(point, esriTextHorizontalAlignment.esriTHARight, esriTextVerticalAlignment.esriTVACenter);
                    graphContainer.AddElement(road as IElement, 0);
                }
                if (info.nz == DkInfo.ROAD_ID)
                {
                    point = GetTextPoint(feature, 2);
                    road = CreateTextElement(point, esriTextHorizontalAlignment.esriTHACenter, esriTextVerticalAlignment.esriTVATop);
                    graphContainer.AddElement(road as IElement, 0);
                }
                if (info.bz == DkInfo.ROAD_ID)
                {
                    point = GetTextPoint(feature, 3);
                    road = CreateTextElement(point, esriTextHorizontalAlignment.esriTHACenter, esriTextVerticalAlignment.esriTVABottom);
                    graphContainer.AddElement(road as IElement, 0);
                }
            }
            return true;
        }
        private  ITextElement CreateTextElement(IPoint point, esriTextHorizontalAlignment h, esriTextVerticalAlignment v)
        {
            ITextElement element = new TextElementClass();

            System.Drawing.Font font = new System.Drawing.Font("Arial", 8, System.Drawing.FontStyle.Bold);
            ITextSymbol textSymbol = new TextSymbolClass();
            textSymbol.Size = 2;
            IRgbColor color = new RgbColorClass();
            color.RGB = 0;
            textSymbol.Color = color;
            textSymbol.Font = (stdole.IFontDisp)OLE.GetIFontDispFromFont(font);
            textSymbol.HorizontalAlignment = h;
            textSymbol.VerticalAlignment = v;
            textSymbol.Text = "路";
            element.ScaleText = true;
            element.Symbol = textSymbol;
            element.Text = "路";

            IElement geoElement = element as IElement;
            geoElement.Geometry = point;

            return element;
        }
        /// <summary>
        /// 根据方向和地块多边形确定标签位置
        /// </summary>
        /// <param name="target">地块多边形</param>
        /// <param name="dir">方向</param>
        /// <returns></returns>
        private  IPoint GetTextPoint(IFeature target, int dir)
        {
            IEnvelope queryEnvelop = target.Shape.Envelope;
            queryEnvelop.Expand(1.1, 1.3, true);
            IEnvelope envelope = target.Shape.Envelope;

            IPoint A = null;
            IPoint B = null;
            IPoint C = new PointClass();
            C.SpatialReference = envelope.SpatialReference;
            C.PutCoords((queryEnvelop.XMax + queryEnvelop.XMin) / 2, (queryEnvelop.YMax + queryEnvelop.YMin) / 2);

            switch (dir)
            {
                case 0:     //东至
                    A = queryEnvelop.UpperRight;
                    B = queryEnvelop.LowerRight;
                    break;
                case 1:     //西至
                    A = queryEnvelop.UpperLeft;
                    B = queryEnvelop.LowerLeft;
                    break;
                case 2:     //南至
                    A = queryEnvelop.LowerLeft;
                    B = queryEnvelop.LowerRight;
                    break;
                case 3:     //北至
                    A = queryEnvelop.UpperRight;
                    B = queryEnvelop.UpperLeft;
                    break;
                default:    //错误
                    return null;
            }
            ILine line = null;
            IRing ring = null;
            ISegmentCollection segColl = null;
            segColl = new RingClass();
            line = new LineClass();
            line.PutCoords(A, B);
            object missing = Type.Missing;
            segColl.AddSegment(line as ISegment, ref missing, ref missing);
            line = new LineClass();
            line.PutCoords(B, C);
            segColl.AddSegment(line as ISegment, ref missing, ref missing);
            ring = segColl as IRing;
            ring.Close();
            IGeometryCollection polygon;
            polygon = new PolygonClass();
            polygon.AddGeometry(ring, ref missing, ref missing);

            IRelationalOperator relate = polygon as IRelationalOperator;

            IPointCollection pointColl = target.ShapeCopy as IPointCollection;

            double avg = 0;
            double cnt = 0;
            for (int i = 0; i < pointColl.PointCount; ++i)
            {
                IPoint tmp = pointColl.get_Point(i);
                if (relate.Contains(pointColl.get_Point(i)))
                {
                    switch (dir)
                    {
                        case 0:
                        case 1:
                            avg += pointColl.get_Point(i).Y;
                            ++cnt;
                            break;
                        case 2:
                        case 3:
                            avg += pointColl.get_Point(i).X;
                            ++cnt;
                            break;
                        default:
                            return null;
                    }
                }
            }
            IPoint ret = new PointClass();
            ret.SpatialReference = envelope.SpatialReference;
            if (cnt == 0)
            {
                if (dir == 0 || dir == 1)
                    avg = C.Y;
                if (dir == 2 || dir == 3)
                    avg = C.X;
            }
            else
                avg /= cnt;
            switch (dir)
            {
                case 0:
                    ret.PutCoords(envelope.XMax, avg);
                    break;
                case 1:
                    ret.PutCoords(envelope.XMin, avg);
                    break;
                case 2:
                    ret.PutCoords(avg, envelope.YMin);
                    break;
                case 3:
                    ret.PutCoords(avg, envelope.YMax);
                    break;
                default:
                    return null;
            }

            return ret;
        }

        private static bool InsertPointClass(IFeatureClass zdt, List<DkInfo> dkInfoList, IFeatureClass pointFC)
        {
            IFeature feature = null;
            IFeatureCursor pointCursor = pointFC.Insert(true);
            IFeatureBuffer pointBuf = null;
            IPoint point = null;
            try
            {
                foreach (DkInfo info in dkInfoList)
                {
                    feature = zdt.GetFeature(info.dkid);
                    point = new PointClass();
                    point.SpatialReference = feature.ShapeCopy.SpatialReference;
                    point.X = (feature.Shape.Envelope.XMax + feature.Shape.Envelope.XMin) / 2;
                    point.Y = (feature.Shape.Envelope.YMax + feature.Shape.Envelope.YMin) / 2;
                    pointBuf = pointFC.CreateFeatureBuffer();
                    pointBuf.Shape = point;

                    pointCursor.InsertFeature(pointBuf);
                }

                pointCursor.Flush();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                return false;
            }
            return true;
        }
        private  void AddPostScript(IWorkbook workbook, int sheetIndex, string cbfbm, string cbfmc,
            string familyCount, int fieldCount, double scmj, string address, int currentPage, int totalPage)
        {
            ISheet sheet = workbook.GetSheetAt(sheetIndex);
            NPOI.SS.UserModel.IRow row = null;
            row = sheet.GetRow(6);
        }

        /// <summary>
        /// 以多边形的外包矩形为判断依据
        /// </summary>
        /// <param name="centerPolygon">中心多边形</param>
        /// <param name="targetPolygon">目标多边形</param>
        /// <returns>索引值</returns>
        private  int GetDirection(IEnvelope centerPolygon, IEnvelope targetPolygon)
        {
            //获取中心点和目标中心点
            IPoint centerPoint = new PointClass();
            centerPoint.PutCoords((centerPolygon.XMax + centerPolygon.XMin) / 2, (centerPolygon.YMax + centerPolygon.YMin) / 2);
            IPoint targetPoint = new PointClass();
            targetPoint.PutCoords((targetPolygon.XMax + targetPolygon.XMin) / 2, (targetPolygon.YMax + targetPolygon.YMin) / 2);
            //判断方位
            IPoint vectorPoint = new PointClass();
            vectorPoint.PutCoords(targetPoint.X - centerPoint.X, targetPoint.Y - centerPoint.Y);
            //MessageBox.Show(string.Format("x:{0};y:{1}", vectorPoint.X, vectorPoint.Y));
            //如果方向是东
            if (vectorPoint.X > 0 && (vectorPoint.X >= Math.Abs(vectorPoint.Y)))
            {
                return 0;
            }
            //如果方向是西
            if (vectorPoint.X < 0 && (Math.Abs(vectorPoint.X)) >= (Math.Abs(vectorPoint.Y)))
            {
                return 1;
            }
            //如果方向是南
            if (vectorPoint.Y < 0 && (Math.Abs(vectorPoint.Y)) > (Math.Abs(vectorPoint.X)))
            {
                return 2;
            }
            //如果方向是北
            if (vectorPoint.Y > 0 && (Math.Abs(vectorPoint.Y)) > (Math.Abs(vectorPoint.X)))
            {
                return 3;
            }
            return -1;
        }
        /// <summary>
        /// 根据承包方编码，来创建只包含承包方地块要素和与承包方地块相邻地块的要素的要素类
        /// </summary>
        /// <param name="zdt">宗地图要素类</param>
        /// <param name="cbfbm">承包方编码</param>
        /// <param name="dkInfoList">(out)地块信息列表</param>
        /// <param name="dkPointFC">(out)输出的地块点要素</param>
        /// <param name="outFC">(out)输出的地块要素</param>
        /// <param name="dbUrl">输出临时要素类的mdb数据库</param>
        /// <returns></returns>
        private  bool BuildNewFeatureClass(IFeatureClass zdt, string cbfbm, out IFeatureClass outFC, out IFeatureClass dkPointFC, out List<DkInfo> dkInfoList, string dbUrl)
        {
            outFC = null;
            dkInfoList = null;
            dkPointFC = null;
            if (zdt == null || cbfbm == null || dbUrl == null)
                return false;

            dkInfoList = new List<DkInfo>();
            IQueryFilter query = new QueryFilterClass();
            query.WhereClause = "CBFBM = \"" + cbfbm + "\"";
            IFeatureCursor cursor = zdt.Search(query, false);
            IFeature feature = null;
            bool ret = true;
            DkInfo dkInfo = null;
            int dzIndex = zdt.Fields.FindField("DKDZ");
            int xzIndex = zdt.Fields.FindField("DKXZ");
            int nzIndex = zdt.Fields.FindField("DKNZ");
            int bzIndex = zdt.Fields.FindField("DKBZ");
            int cbfmcIndex = zdt.Fields.FindField("CBFMC");

            ISpatialFilter spatialQuery = new SpatialFilterClass();
            spatialQuery.GeometryField = zdt.ShapeFieldName;
            spatialQuery.SpatialRel = esriSpatialRelEnum.esriSpatialRelTouches;

            string dkdz = null;
            string dkxz = null;
            string dknz = null;
            string dkbz = null;
            string tmpCbfmc = null;

            while ((feature = cursor.NextFeature()) != null)
            {
                spatialQuery.Geometry = feature.ShapeCopy;
                dkInfo = new DkInfo();
                dkInfo.dkid = feature.OID;
                dkdz = feature.get_Value(dzIndex).ToString().Trim();
                dkxz = feature.get_Value(xzIndex).ToString().Trim();
                dknz = feature.get_Value(nzIndex).ToString().Trim();
                dkbz = feature.get_Value(bzIndex).ToString().Trim();

                IFeatureCursor scursor = zdt.Search(spatialQuery, false);
                IFeature tmp = null;
                int tmpIndex = -1;
                while ((tmp = scursor.NextFeature()) != null)
                {
                    tmpCbfmc = tmp.get_Value(cbfmcIndex).ToString().Trim();
                    tmpIndex = GetDirection(feature.Shape.Envelope, tmp.Shape.Envelope);
                    if (tmpCbfmc == dkdz && tmpIndex == 0)      //东至
                    {
                        dkInfo.dz = tmp.OID;
                    }
                    else if (tmpCbfmc == dkxz && tmpIndex == 1)  //西至
                    {
                        dkInfo.xz = tmp.OID;
                    }
                    else if (tmpCbfmc == dknz && tmpIndex == 2)   //南至
                    {
                        dkInfo.nz = tmp.OID;
                    }
                    else if (tmpCbfmc == dkbz && tmpIndex == 3)   //北至
                    {
                        dkInfo.bz = tmp.OID;
                    }
                }
                bool valid = true;
                if (dkdz != "路" && dkInfo.dz == DkInfo.ROAD_ID)
                {
                    valid = false;
                }
                else if (dkxz != "路" && dkInfo.xz == DkInfo.ROAD_ID)
                {
                    valid = false;
                }
                else if (dknz != "路" && dkInfo.nz == DkInfo.ROAD_ID)
                {
                    valid = false;
                }
                else if (dkbz != "路" && dkInfo.bz == DkInfo.ROAD_ID)
                {
                    valid = false;
                }

                if (!valid)
                {
                    //System.Windows.Forms.MessageBox.Show("重新匹配东西南北至错误！");
                }
                dkInfoList.Add(dkInfo);
            }

            HashSet<int> idSet = new HashSet<int>();
            foreach (DkInfo info in dkInfoList)
            {
                idSet.Add(info.dkid);
                idSet.Add(info.dz);
                idSet.Add(info.nz);
                idSet.Add(info.xz);
                idSet.Add(info.bz);
            }
            idSet.Remove(-1);

            IFeatureCursor tmpCursor = zdt.GetFeatures(idSet.ToArray(), false);
            IAeFactory pAeFactory=new PersonalGeoDatabase(dbUrl);

            IFeatureWorkspace inmemWor = pAeFactory.OpenFeatrueWorkspace();
            pAeFactory.DeleteIfExist("tmp_zdt");
            pAeFactory.DeleteIfExist("point");
            //delIfExist(inmemWor, esriDatasetType.esriDTFeatureClass, "tmp_zdt");
            //delIfExist(inmemWor, esriDatasetType.esriDTFeatureClass, "point");

            outFC = CloneFeatureClassInWorkspace(zdt, inmemWor, "tmp_zdt", null);

            IFieldsEdit fieldsEdit = new FieldsClass() as IFieldsEdit;
            IGeometryDef geometryDef = new GeometryDefClass();
            IGeometryDefEdit geometryDefEdit = geometryDef as IGeometryDefEdit;
            geometryDefEdit.GeometryType_2 = esriGeometryType.esriGeometryPoint;
            geometryDefEdit.GridCount_2 = 1;
            geometryDefEdit.set_GridSize(0, 10);
            geometryDefEdit.SpatialReference_2 = (zdt as IGeoDataset).SpatialReference;

            IField fieldShape = new FieldClass();
            IFieldEdit fieldShapeEdit = fieldShape as IFieldEdit;
            fieldShapeEdit.Name_2 = "SHAPE";
            fieldShapeEdit.AliasName_2 = "SHAPE";
            fieldShapeEdit.Type_2 = esriFieldType.esriFieldTypeGeometry;
            fieldShapeEdit.GeometryDef_2 = geometryDef;
            fieldsEdit.AddField(fieldShape);

            IField fieldOID = new FieldClass();
            IFieldEdit fieldOIDEdit = fieldOID as IFieldEdit;
            fieldOIDEdit.Name_2 = "OBJECTID";
            fieldOIDEdit.AliasName_2 = "OBJECTID";
            fieldOIDEdit.Type_2 = esriFieldType.esriFieldTypeOID;
            fieldsEdit.AddField(fieldOID);

            dkPointFC = inmemWor.CreateFeatureClass("point", fieldsEdit, null, null, esriFeatureType.esriFTSimple, "SHAPE", "");


            IFeatureCursor outCursor = outFC.Insert(true);

            IFeature temp = null;
            IFeatureBuffer buffer = null;

            IFields fields = outFC.Fields;
            Hashtable hashmap = new Hashtable();
            hashmap.Add(-1, -1);
            int oIndex = outFC.FindField(outFC.OIDFieldName);
            int sIndex = outFC.FindField(outFC.ShapeFieldName);
            int tmpOid = -1;

            Hashtable indexMap = new Hashtable();
            for (int i = 0; i < outFC.Fields.FieldCount; ++i)
            {
                indexMap.Add(i, zdt.FindField(outFC.Fields.get_Field(i).Name));
            }
            try
            {
                while ((temp = tmpCursor.NextFeature()) != null)
                {
                    buffer = outFC.CreateFeatureBuffer();

                    for (int i = 0; i < fields.FieldCount; ++i)
                    {
                        if (i == oIndex)
                            continue;
                        if (i == sIndex)
                        {
                            buffer.Shape = temp.ShapeCopy;
                            continue;
                        }
                        try
                        {
                            if (fields.get_Field(i).Editable)
                            {
                                buffer.set_Value(i, temp.get_Value((int)indexMap[i]));
                            }
                        }
                        catch
                        {
                        }
                    }
                    tmpOid = (int)outCursor.InsertFeature(buffer);
                    hashmap.Add(temp.OID, tmpOid);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                return false;
            }
            outCursor.Flush();

            foreach (DkInfo info in dkInfoList)
            {
                info.dkid = (int)hashmap[info.dkid];
                info.dz = (int)hashmap[info.dz];
                info.xz = (int)hashmap[info.xz];
                info.nz = (int)hashmap[info.nz];
                info.bz = (int)hashmap[info.bz];
            }
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(inmemWor);
            return ret;
        }
        private  IFeatureClass CloneFeatureClassInWorkspace(IFeatureClass OldFeatureClass, IFeatureWorkspace SaveFeatWorkspace, string FeatClsName, IEnvelope pDomainEnv)
        {
            IFields pFields = CloneFeatureClassFields(OldFeatureClass, null);
            //delIfExist(SaveFeatWorkspace, esriDatasetType.esriDTFeatureClass,FeatClsName);

            var pfeaureclass = SaveFeatWorkspace.CreateFeatureClass(FeatClsName, pFields, null, null, esriFeatureType.esriFTSimple, OldFeatureClass.ShapeFieldName, "");
            // System.Runtime.InteropServices.Marshal.ReleaseComObject(SaveFeatWorkspace);
            return pfeaureclass;
        }
        /// <summary>
        /// 复制要素类字段
        /// </summary>
        /// <param name="pFeatureClass"></param>
        /// <param name="pDomainEnv"></param>
        /// <returns></returns>
        private  IFields CloneFeatureClassFields(IFeatureClass pFeatureClass, IEnvelope pDomainEnv)
        {
            IFields pFields = new FieldsClass();
            IFieldsEdit pFieldsEdit = (IFieldsEdit)pFields;
            //根据传入的要素类,将除了shape字段之外的字段复制  
            long nOldFieldsCount = pFeatureClass.Fields.FieldCount;
            long nOldGeoIndex = pFeatureClass.Fields.FindField(pFeatureClass.ShapeFieldName);
            for (int i = 0; i < nOldFieldsCount; i++)
            {
                if (i != nOldGeoIndex)
                {
                    pFieldsEdit.AddField(pFeatureClass.Fields.get_Field(i));
                }
                else
                {
                    IGeometryDef pGeomDef = new GeometryDefClass();
                    IGeometryDefEdit pGeomDefEdit = (IGeometryDefEdit)pGeomDef;
                    ISpatialReference pSR = null;
                    if (pDomainEnv != null)
                    {
                        pSR = new UnknownCoordinateSystemClass();
                        pSR.SetDomain(pDomainEnv.XMin, pDomainEnv.XMax, pDomainEnv.YMin, pDomainEnv.YMax);
                    }
                    else
                    {
                        IGeoDataset pGeoDataset = pFeatureClass as IGeoDataset;
                        pSR = pGeoDataset.SpatialReference;
                    }
                    //设置新要素类Geometry的参数  
                    pGeomDefEdit.GeometryType_2 = pFeatureClass.ShapeType;
                    pGeomDefEdit.GridCount_2 = 1;
                    pGeomDefEdit.set_GridSize(0, 10);
                    pGeomDefEdit.AvgNumPoints_2 = 2;
                    pGeomDefEdit.SpatialReference_2 = pSR;
                    //产生新的shape字段  
                    IField pField = new FieldClass();
                    IFieldEdit pFieldEdit = (IFieldEdit)pField;
                    pFieldEdit.Name_2 = pFeatureClass.Fields.get_Field(i).Name;
                    pFieldEdit.AliasName_2 = pFeatureClass.Fields.get_Field(i).AliasName;
                    pFieldEdit.Type_2 = esriFieldType.esriFieldTypeGeometry;
                    pFieldEdit.GeometryDef_2 = pGeomDef;
                    pFieldsEdit.AddField(pField);
                }
            }
            return pFields;
        }
        /// <summary>
        /// 修复地图文档中图层的数据源
        /// </summary>
        /// <param name="map">要修复的地图</param>
        /// <param name="zdtFC">宗地图要素类</param>
        /// <param name="bjxFC">边界线要素类</param>
        /// <param name="pointFC">地块点要素类</param>
        /// <returns></returns>
        private static bool FixDataSource(IMap map, IFeatureClass zdtFC, IFeatureClass bjxFC, IFeatureClass pointFC)
        {
            if (map == null || zdtFC == null || bjxFC == null)
                return false;
            ILayer layer = null;
            bool zdt = false, bjx = false, p = false;
            for (int i = 0; i < map.LayerCount; ++i)
            {
                layer = map.get_Layer(i);
                if (layer.Name == "zdt")
                {
                    (layer as IGeoFeatureLayer).FeatureClass = zdtFC;
                    zdt = true;
                }
                else if (layer.Name == "bjx")
                {
                    (layer as IGeoFeatureLayer).FeatureClass = bjxFC;
                    bjx = true;
                }
                else if (layer.Name == "point")
                {
                    (layer as IGeoFeatureLayer).FeatureClass = pointFC;
                    p = true;
                }
                if (bjx && zdt && p)
                    return true;
            }
            return false;
        }

        /// <summary>
        /// 设置要素显示的方式和要素注记的显示方式
        /// </summary>
        /// <param name="map">要修复的地图</param>
        /// <param name="cbfbm">承办方编码</param>
        /// <returns></returns>
        private static bool SetQuery(IMap map, String cbfbm)
        {
            ILayer layer = null;
            IGeoFeatureLayer geoLayer = null;
            for (int i = 0; i < map.LayerCount; ++i)
            {
                layer = map.get_Layer(i);
                if (layer.Name == "zdt")
                {
                    geoLayer = layer as IGeoFeatureLayer;
                    break;
                }
            }

            if (geoLayer == null)
            {
                System.Windows.Forms.MessageBox.Show("未找到宗地图图层！");
                return false;
            }

            IUniqueValueRenderer render = geoLayer.Renderer as IUniqueValueRenderer;

            if (render == null || render.ValueCount != 1)
            {
                System.Windows.Forms.MessageBox.Show("地图文档不正确！");
                return false;
            }
            render.set_Value(0, cbfbm);

            IAnnotateLayerPropertiesCollection IPALPColl = geoLayer.AnnotationProperties;
            IAnnotateLayerProperties layerProp = null;
            IElementCollection element = null;

            if (IPALPColl == null || IPALPColl.Count != 2)
            {
                System.Windows.Forms.MessageBox.Show("地图文档不正确！");
                return false;
            }

            bool d = false, cbf = false;
            for (int i = 0; i < IPALPColl.Count; ++i)
            {
                IPALPColl.QueryItem(i, out layerProp, out element, out element);
                if (layerProp.Class == "Default")
                {
                    layerProp.WhereClause = "[CBFBM] <> \"" + cbfbm + "\"";
                    d = true;
                }
                else if (layerProp.Class == "CBF")
                {
                    layerProp.WhereClause = "[CBFBM] = \"" + cbfbm + "\"";
                    cbf = true;
                }
                if (cbf && d)
                    break;
            }

            if (!(cbf && d))
            {
                System.Windows.Forms.MessageBox.Show("地图文档不正确！");
                return false;
            }

            return true;
        }

        /// <summary>
        /// 输出图片到本地文件
        /// </summary>
        /// <param name="pActiveView"></param>
        /// <param name="envelop">输出的地理位置</param>
        /// <param name="imgWidth">图片宽度</param>
        /// <param name="imgHeight">图片高度</param>
        /// <param name="imagepath">图片路径</param>
        /// <returns></returns>
        public static bool ExportImageToLocal(IActiveView pActiveView, IEnvelope envelop, int imgWidth, int imgHeight, string imagepath)
        {
            IEnvelope pEnvelope = new EnvelopeClass();
            ITrackCancel pTrackCancel = new CancelTrackerClass();

            tagRECT ptagRECT = new tagRECT();// pActiveView.ExportFrame;
            ptagRECT.left = 0;
            ptagRECT.top = 0;
            ptagRECT.right = imgWidth;// (int)pActiveView.Extent.Width;
            ptagRECT.bottom = imgHeight;// (int)pActiveView.Extent.Height;

            int pResolution = (int)(pActiveView.ScreenDisplay.DisplayTransformation.Resolution);
            if (pResolution == 0)
            {
                pResolution = 96;
            }
            pEnvelope.PutCoords(ptagRECT.left, ptagRECT.bottom, ptagRECT.right, ptagRECT.top);

            IEnvelope newEnvelope = envelop.Envelope;

            newEnvelope.Expand(newEnvelope.Width * 0.025, newEnvelope.Height * 0.025, false);

            ExportJPEGClass bitmap = new ExportJPEGClass();
            bitmap.Resolution = pResolution;
            bitmap.ExportFileName = imagepath;
            bitmap.PixelBounds = pEnvelope;

            pActiveView.Output(bitmap.StartExporting(), pResolution, ref ptagRECT, newEnvelope, pTrackCancel);
            bitmap.FinishExporting();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(bitmap);

            return true;
        }
    }
}
