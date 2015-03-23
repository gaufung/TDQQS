using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Aspose.Pdf.Facades;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Geometry;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using TdqqClient.Services.AE;

namespace TdqqClient.Services.Export.ExportSingle
{
    class DkExport:ExportBase
    {

        public DkExport(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        { }

        public override void Export(string cbfmc, string cbfbm, string folderPath, string edgeFeature = "")
        {

            var fields = Fields(cbfbm);
            var jzxFeature = SelectFeature + "_JZX";
            var jxdFeature = SelectFeature + "_JZD";
            //每个人按名字一行hashrow
            var bmIdHashTable = new Hashtable();
            var bmXmHashTable = new Hashtable();
            var idBmHashTable = new Hashtable();
            for (int i = 0; i < fields.Count; i++)
            {
                string tempDKOBJECT = fields[i].ObjectId + "";
                string tempCbfmc = fields[i].Cbfmc + "";
                string tempCbfbm = fields[i].Cbfbm + "";
                string tempDkbm = fields[i].Dkbm + "";
                //Hashtable索引中是不是已经加载过CBFBM
                if (bmIdHashTable.Contains(tempCbfbm) == false)
                    bmIdHashTable.Add(tempCbfbm, tempDKOBJECT);
                else
                    bmIdHashTable[tempCbfbm] = bmIdHashTable[tempCbfbm].ToString() + ',' + tempDKOBJECT;

                if (bmXmHashTable.Contains(tempCbfbm) == false)
                    bmXmHashTable.Add(tempCbfbm, tempCbfmc);

                if (idBmHashTable.Contains(tempDKOBJECT) == false)
                    idBmHashTable.Add(tempDKOBJECT, tempDkbm);
            }
            int tempNum = 0;
            var ztdkObjects = new List<int>();
            foreach (DictionaryEntry de in bmIdHashTable)
            {

                string tempCbfbm = de.Key.ToString();
                //拿出承包方调查员日期等参数
                Hashtable infoHashTable = Getinfo(tempCbfbm);

                string tempObjectIDs = de.Value.ToString();
                string[] objectIds = tempObjectIDs.Split(',');
                var fids = new List<int>();
                foreach (string s in objectIds)
                {
                    fids.Add(Convert.ToInt32(s));
                }
                string singleFolderPath = folderPath;
                string excelUrl;
                var twofids = new List<int>();
                for (int i = 0; i < fids.Count; i++)
                {
                    int jzNum = GetJZNum(PersonDatabase, fids[i], SelectFeature, jzxFeature, jxdFeature);
                    if (jzNum == -1)
                    {
                        throw new InvalidDataException("请做一次拓扑检查");
                    }
                    if (jzNum > 9)
                    {
                        //新的把点分开，写到几个Excel
                        excelUrl = singleFolderPath + @"\" + tempCbfbm.Substring(14) + '_' + bmXmHashTable[tempCbfbm] + "_02承包地块调查表" + '-' + fids[i] + ".xls";
                        CreateManyCTable(PersonDatabase, fids[i], excelUrl, SelectFeature, jzxFeature, jxdFeature, infoHashTable);
                        Export2Pdf.Excel2Pdf(excelUrl);
                        ztdkObjects.Add(fids[i]);
                    }
                    else
                    {
                        twofids.Add(fids[i]);
                        if (twofids.Count == 2)
                        {
                            int[] tempArr = new int[2];
                            tempArr[0] = twofids[0];
                            tempArr[1] = twofids[1];
                            excelUrl = singleFolderPath + @"\" + tempCbfbm.Substring(14) + '_' + bmXmHashTable[tempCbfbm] + "_02承包地块调查表" + '-' + tempArr[0] + '-' + tempArr[1] + ".xls";
                            CreateOneC2Table(PersonDatabase, tempArr, excelUrl, SelectFeature, jzxFeature, jxdFeature, infoHashTable);
                            Export2Pdf.Excel2Pdf(excelUrl);
                            twofids.Clear();
                        }
                    }
                }
                foreach (int f in twofids)
                {
                    excelUrl = singleFolderPath + @"\" + tempCbfbm.Substring(14) + '_' + bmXmHashTable[tempCbfbm] + "_02承包地块调查表" + '-' + f + ".xls";

                    CreateOneCTable(PersonDatabase, f, excelUrl, SelectFeature, jzxFeature, jxdFeature, infoHashTable);
                    Export2Pdf.Excel2Pdf(excelUrl);
                }
                tempNum = tempNum + 1;
            }
            Concatenate(folderPath, cbfmc, cbfbm);
            DeleteFiles(folderPath);
        }

        /// <summary>
        /// 将某个农户所有的地块调查表合并成一个
        /// </summary>
        /// <param name="folderPath">文件夹位置</param>
        /// <param name="cbfmc">承包方名称</param>
        /// <param name="cbfbm">承包方编码</param>
        private void Concatenate(string folderPath, string cbfmc, string cbfbm)
        {

            var targetPdfPath = folderPath + @"\" + cbfbm.Substring(14) + "_" + cbfmc + "_" + "02承包地块调查表.pdf";
            var inFileStream = GetToMergePdf(folderPath, cbfbm);
            var outFileStream = new FileStream(targetPdfPath, FileMode.Create);
            var pdfEditor = new PdfFileEditor();
            pdfEditor.Concatenate(inFileStream.ToArray(), outFileStream);
            outFileStream.Close();
            foreach (var stream in inFileStream)
            {
                stream.Close();
            }
        }



        /// <summary>
        /// 获取该农户地块调查表
        /// </summary>
        /// <param name="folderPath">文件夹路径</param>
        /// <param name="cbfbm">承包方编码</param>
        /// <returns>文件流对象</returns>
        private List<Stream> GetToMergePdf(string folderPath, string cbfbm)
        {
            var dir = new DirectoryInfo(folderPath);
            var listFile = new List<Stream>();
            foreach (var dChild in dir.GetFiles("*.pdf"))
            {
                var name = dChild.Name;
                if (name.StartsWith(cbfbm.Substring(14)))
                {
                    listFile.Add(new FileStream(dChild.FullName, FileMode.Open));
                }
            }
            return listFile;
        }

        private Hashtable Getinfo(string cbfbm)
        {
            var tempHashTable = new Hashtable();
            var dcsh = DcSh();
            string cbfdcy = dcsh.Cbfdcy;
            tempHashTable.Add("dcy", cbfdcy);
            var cbfdcrq = dcsh.Cbfdcrq;            
            tempHashTable.Add("dcrq", cbfdcrq);
            var fbf = Fbf();
            string fbffzr = fbf == null ? string.Empty : fbf.Fbffzrxm;
            tempHashTable.Add("fbffzr", fbffzr);
            return tempHashTable;
        }
        private void DeleteFiles(string folderPath)
        {
            var files = GetToDeleteFile(folderPath);
            foreach (var file in files)
            {
                File.Delete(file);
            }

        }

        private IEnumerable<string> GetToDeleteFile(string folderPath)
        {
            var dir = new DirectoryInfo(folderPath);
            foreach (var file in dir.GetFiles())
            {
                var name = file.FullName;
                if (!name.EndsWith("02承包地块调查表.pdf"))
                {
                    yield return name;
                }
            }
        }
        private int GetJZNum(string databaseUrl, int fid, string zdtname, string jzxname, string jzdname)
        {
            IAeFactory pAeFactory = new PersonalGeoDatabase(databaseUrl);
            IFeatureWorkspace workspace = pAeFactory.OpenFeatrueWorkspace();
            IFeatureClass zdt = workspace.OpenFeatureClass(zdtname);
            IFeatureClass jzx = workspace.OpenFeatureClass(jzxname);
            IFeatureClass jzd = workspace.OpenFeatureClass(jzdname);
            IFeature zdtF = zdt.GetFeature(fid);
            //IFeatureWorkspace workspace, IFeature zdtF, IFeatureClass jzx, IFeatureClass jzd
            ISpatialFilter sf = new SpatialFilterClass();
            sf.Geometry = zdtF.ShapeCopy;
            sf.GeometryField = "SHAPE";
            sf.SpatialRel = esriSpatialRelEnum.esriSpatialRelRelation;
            sf.SpatialRelDescription = "F*TT*TF*T";
            var cursor = jzd.Search(sf, false);
            IFeature tmpFeature = null;
            List<IFeature> jzdList = new List<IFeature>();
            while ((tmpFeature = cursor.NextFeature()) != null)
            {
                jzdList.Add(tmpFeature);
            }
            //筛选界址线
            sf = new SpatialFilterClass();
            sf.GeometryField = "SHAPE";
            sf.Geometry = zdtF.ShapeCopy;
            sf.SpatialRel = esriSpatialRelEnum.esriSpatialRelRelation;
            sf.SpatialRelDescription = "FFTTT*FF*";

            cursor = jzx.Search(sf, false);
            tmpFeature = null;
            List<IFeature> jzxList = new List<IFeature>();

            while ((tmpFeature = cursor.NextFeature()) != null)
            {
                jzxList.Add(tmpFeature);
            }
            if (jzdList.Count > jzxList.Count)
                return -1;
            else
                return jzdList.Count;
        }

        private bool CreateManyCTable(string databaseUrl, int fid, string outpath, string zdtname, string jzxname, string jzdname, Hashtable InfoHashTable)
        {
            try
            {
                IAeFactory pAeFactory = new PersonalGeoDatabase(databaseUrl);
                IFeatureWorkspace workspace = pAeFactory.OpenFeatrueWorkspace();
                IFeatureClass zdt = workspace.OpenFeatureClass(zdtname);
                IFeatureClass jzx = workspace.OpenFeatureClass(jzxname);
                IFeatureClass jzd = workspace.OpenFeatureClass(jzdname);
                Hashtable tempHashTable = InfoHashTable;
                var ret = CreateCManyTable(workspace, outpath, zdt.GetFeature(fid), jzx, jzd, tempHashTable);
                Marshal.FinalReleaseComObject(zdt);
                Marshal.FinalReleaseComObject(workspace);
                GC.WaitForPendingFinalizers();
                GC.Collect();
                return ret;
            }
            catch
            {
            }
            return false;
        }
        private bool CreateCManyTable(IFeatureWorkspace workspace, string outpath, IFeature zdtF, IFeatureClass jzx, IFeatureClass jzd, Hashtable tempHashTable)
        {
            Hashtable InfoHashTable = tempHashTable;
            List<IFeature> jzdSorted = new List<IFeature>();
            List<IFeature> jzxSorted = new List<IFeature>();
            // 拿出地块所有界址点界址线
            ISpatialFilter sf = new SpatialFilterClass();
            sf.Geometry = zdtF.ShapeCopy;
            sf.GeometryField = "SHAPE";
            sf.SpatialRel = esriSpatialRelEnum.esriSpatialRelRelation;
            sf.SpatialRelDescription = "F*TT*TF*T";
            var cursor = jzd.Search(sf, false);
            IFeature tmpFeature = null;
            List<IFeature> jzdList = new List<IFeature>();
            while ((tmpFeature = cursor.NextFeature()) != null)
            {
                jzdList.Add(tmpFeature);
            }
            IGeometryCollection gc = zdtF.ShapeCopy as IGeometryCollection;
            if (gc.GeometryCount > 1)
            {
                int a = gc.GeometryCount;
            }
            sf = new SpatialFilterClass();
            sf.GeometryField = "SHAPE";
            sf.Geometry = zdtF.ShapeCopy;
            sf.SpatialRel = esriSpatialRelEnum.esriSpatialRelRelation;
            sf.SpatialRelDescription = "FFTTT*FF*";

            cursor = jzx.Search(sf, false);
            tmpFeature = null;
            List<IFeature> jzxList = new List<IFeature>();

            while ((tmpFeature = cursor.NextFeature()) != null)
            {
                jzxList.Add(tmpFeature);
            }
            if (jzdList.Count != jzxList.Count)
            {
                //  System.Windows.Forms.MessageBox.Show("筛选出的界址点与界址线数量不同！");
                return false;
            }

            int j = jzdList.Count;
            IFeature tmppoint = jzdList[0];
            IFeature tmpline = null;
            int gcCount = (zdtF.ShapeCopy as IGeometryCollection).GeometryCount - 1;
            while (jzxList.Count > 0)
            {
                if (tmppoint == null)
                {
                    if (gcCount > 0)
                    {
                        jzxSorted.Add(null);
                        jzdSorted.Add(null);
                        tmppoint = jzdList[0];
                        --gcCount;
                    }
                    else
                    {
                        //   System.Windows.Forms.MessageBox.Show("筛选出的界址点与界址线错误！");
                        return false;
                    }
                }

                tmpline = getRelationLine(tmppoint.ShapeCopy, jzxList);
                if (tmpline == null)
                {

                    //   System.Windows.Forms.MessageBox.Show("筛选出的界址点与界址线错误！");
                    return false;
                }

                jzdSorted.Add(tmppoint);
                jzdList.Remove(tmppoint);
                jzxSorted.Add(tmpline);
                jzxList.Remove(tmpline);

                tmppoint = getRelationPoint(tmpline.ShapeCopy, jzdList);
            }


            //把界址点按照小于等于18个分开
            List<IFeature> mjzdSorted = new List<IFeature>();
            List<IFeature> mjzxSorted = new List<IFeature>();
            //记录为整18的个数
            int eNum = 0;
            for (int i = 0; i < jzdSorted.Count; i++)
            {
                mjzdSorted.Add(jzdSorted[i]);
                mjzxSorted.Add(jzxSorted[i]);
                {
                    if (mjzdSorted.Count == 18)
                    {
                        //重命名输出文件名Outpath
                        string[] paths = outpath.Split('.');
                        eNum = eNum + 1;
                        paths[0] = paths[0] + '_' + eNum;
                        string temppath = paths[0] + '.' + paths[1];

                        //创建左右为同一家人的地块
                        CreateCM11Table(workspace, temppath, zdtF, jzx, jzd, mjzdSorted, mjzxSorted, InfoHashTable);
                        mjzdSorted.Clear();
                        mjzxSorted.Clear();
                    }
                }

            }
            //循环分配完了，剩余总数小于18的部分报表
            //总数小于等于9，仅填写左边的表，否则还是要调用填写左右两边表的函数
            if (mjzdSorted.Count != 0 && mjzdSorted.Count <= 9)
            {
                //(IFeatureWorkspace workspace, string outpath, IFeature zdtF, IFeatureClass jzx, IFeatureClass jzd)
                //CreateCM1Table(IFeatureWorkspace workspace, string outpath, IFeature zdtF, List<IFeature> pointjzdsorted, List<IFeature> linejzdsorted)
                string[] paths = outpath.Split('.');
                eNum = eNum + 1;
                paths[0] = paths[0] + '_' + eNum;
                string temppath = paths[0] + '.' + paths[1];
                CreateCM1Table(workspace, temppath, zdtF, jzx, jzd, mjzdSorted, mjzxSorted, InfoHashTable);
            }
            if (mjzdSorted.Count != 0 && mjzdSorted.Count > 9)
            {
                string[] paths = outpath.Split('.');
                eNum = eNum + 1;
                paths[0] = paths[0] + '_' + eNum;
                string temppath = paths[0] + '.' + paths[1];
                CreateCM11Table(workspace, temppath, zdtF, jzx, jzd, mjzdSorted, mjzxSorted, InfoHashTable);
            }
            return true;
        }

        private bool CreateCM1Table(IFeatureWorkspace workspace, string outpath, IFeature zdtF, IFeatureClass jzx, IFeatureClass jzd, List<IFeature> pointjzdsorted, List<IFeature> linejzdsorted, Hashtable tempHashTable)
        {
            Hashtable InfoHashTable = tempHashTable;
            Hashtable tdytHashTable = new Hashtable();
            tdytHashTable.Add("1", "种植业");
            tdytHashTable.Add("2", "林业");
            tdytHashTable.Add("3", "畜牧业");
            tdytHashTable.Add("4", "渔业");
            tdytHashTable.Add("5", "其他");

            Hashtable tdlylxHashTable = new Hashtable();
            tdlylxHashTable.Add("011", "水田");
            tdlylxHashTable.Add("012", "水浇地");
            tdlylxHashTable.Add("013", "旱地");

            Hashtable dldjHashTable = new Hashtable();
            dldjHashTable.Add("01", "一等地");
            dldjHashTable.Add("02", "二等地");
            dldjHashTable.Add("03", "三等地");
            dldjHashTable.Add("04", "四等地");
            dldjHashTable.Add("05", "五等地");
            dldjHashTable.Add("06", "六等地");
            dldjHashTable.Add("07", "七等地");

            Hashtable jblxHT = new Hashtable();
            jblxHT.Add("1", 1);
            jblxHT.Add("2", 2);
            jblxHT.Add("3", 3);
            jblxHT.Add("4", 4);
            jblxHT.Add("null", 5);

            Hashtable jzxlxHT = new Hashtable();
            jzxlxHT.Add("01", 6);
            jzxlxHT.Add("02", 7);
            jzxlxHT.Add("03", 8);
            jzxlxHT.Add("04", 9);
            jzxlxHT.Add("05", 10);
            jzxlxHT.Add("06", 11);
            jzxlxHT.Add("07", 12);
            jzxlxHT.Add("08", 13);
            jzxlxHT.Add("09", 14);

            Hashtable jzxwzHT = new Hashtable();
            jzxwzHT.Add("1", 15);
            jzxwzHT.Add("2", 16);
            jzxwzHT.Add("3", 17);



            string templateUrl = AppDomain.CurrentDomain.BaseDirectory + @"template\承包地块调查表.xls";


            int fbfbmIndex = zdtF.Fields.FindField("FBFBM");
            int cbfmcIndex = zdtF.Fields.FindField("CBFMC");
            int dkbhIndex = zdtF.Fields.FindField("DKBM");
            int dkmcIndex = zdtF.Fields.FindField("DKMC");
            int htmjIndex = zdtF.Fields.FindField("HTMJ");
            int dzIndex = zdtF.Fields.FindField("DKDZ");
            int xzIndex = zdtF.Fields.FindField("DKXZ");
            int nzIndex = zdtF.Fields.FindField("DKNZ");
            int bzIndex = zdtF.Fields.FindField("DKBZ");
            int tdytIndex = zdtF.Fields.FindField("TDYT");
            int tdlylxIndex = zdtF.Fields.FindField("TDLYLX");
            int dldjIndex = zdtF.Fields.FindField("DLDJ");
            int sfjbntIndex = zdtF.Fields.FindField("SFJBNT");
            int jzdlxIndex = jzd.Fields.FindField("JBLX");
            int jzdhIndex = jzd.Fields.FindField("JZDH");
            int jzxlbIndex = jzx.Fields.FindField("JZXLB");
            int jzxwzIndex = jzx.Fields.FindField("JZXWZ");
            int jzxsmIndex = jzx.Fields.FindField("JZXSM");
            int pldwqlrIndex = jzx.Fields.FindField("PLDWQLR");
            int pldwzjrIndex = jzx.Fields.FindField("PLDWZJR");

            // if (File.Exists(outpath)) File.Delete(outpath);
            File.Copy(templateUrl, outpath, true);
            using (System.IO.FileStream fileStream = new System.IO.FileStream(outpath, FileMode.Open, FileAccess.ReadWrite))
            {
                HSSFWorkbook workbookSource = new HSSFWorkbook(fileStream);
                HSSFSheet sheetSource = (HSSFSheet)workbookSource.GetSheetAt(0);
                //设定合并单元格的样式
                HSSFCellStyle style = (HSSFCellStyle)workbookSource.CreateCellStyle();
                style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.CENTER;
                style.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.CENTER;
                style.BorderBottom = NPOI.SS.UserModel.BorderStyle.THIN;
                style.BorderRight = NPOI.SS.UserModel.BorderStyle.THIN;
                style.BorderLeft = NPOI.SS.UserModel.BorderStyle.THIN;
                style.BorderTop = NPOI.SS.UserModel.BorderStyle.THIN;
                style.WrapText = true;

                string cbfmc = null;

                var tmprow = sheetSource.GetRow(2);
                var tmpcell = tmprow.GetCell(4);
                tmpcell.SetCellValue(zdtF.get_Value(fbfbmIndex).ToString());

                tmpcell = tmprow.GetCell(12);
                cbfmc = zdtF.get_Value(cbfmcIndex).ToString();
                tmpcell.SetCellValue(zdtF.get_Value(cbfmcIndex).ToString());

                tmprow = sheetSource.GetRow(3);
                tmpcell = tmprow.GetCell(4);
                tmpcell.SetCellValue(zdtF.get_Value(dkbhIndex).ToString().Substring(14));

                tmpcell = tmprow.GetCell(12);
                tmpcell.SetCellValue(zdtF.get_Value(dkmcIndex).ToString());

                tmpcell = tmprow.GetCell(20);
                tmpcell.SetCellValue(((double)zdtF.get_Value(htmjIndex)).ToString("f"));

                tmprow = sheetSource.GetRow(4);
                tmpcell = tmprow.GetCell(1);
                tmpcell.SetCellValue(zdtF.get_Value(dzIndex).ToString());

                tmpcell = tmprow.GetCell(9);
                tmpcell.SetCellValue(zdtF.get_Value(nzIndex).ToString());

                tmpcell = tmprow.GetCell(16);
                tmpcell.SetCellValue(zdtF.get_Value(xzIndex).ToString());

                tmpcell = tmprow.GetCell(20);
                tmpcell.SetCellValue(zdtF.get_Value(bzIndex).ToString());

                tmprow = sheetSource.GetRow(5);
                tmpcell = tmprow.GetCell(1);
                string tmpstr = tdytHashTable[zdtF.get_Value(tdytIndex)].ToString();
                IRichTextString rich = tmpcell.RichStringCellValue;
                IFont font = workbookSource.GetFontAt(rich.GetFontAtIndex(rich.String.Length - 1));
                tmpstr = rich.String.Replace("□" + tmpstr, "■" + tmpstr);
                rich = new HSSFRichTextString(tmpstr);
                rich.ApplyFont(tmpstr.IndexOf('_'), tmpstr.LastIndexOf('_'), font);
                tmpcell.SetCellValue(rich);

                tmpcell = tmprow.GetCell(15);
                tmpstr = tdlylxHashTable[zdtF.get_Value(tdlylxIndex)].ToString();
                rich = tmpcell.RichStringCellValue;
                font = workbookSource.GetFontAt(rich.GetFontAtIndex(rich.String.Length - 1));
                tmpstr = rich.String.Replace("□" + tmpstr, "■" + tmpstr);
                rich = new HSSFRichTextString(tmpstr);
                rich.ApplyFont(tmpstr.IndexOf('_'), tmpstr.LastIndexOf('_'), font);
                tmpcell.SetCellValue(rich);

                tmpcell = tmprow.GetCell(10);
                tmpstr = dldjHashTable[zdtF.get_Value(dldjIndex)].ToString();
                tmpcell.SetCellValue(tmpstr);

                tmpcell = tmprow.GetCell(21);
                tmpstr = zdtF.get_Value(sfjbntIndex).ToString();
                if ("2".CompareTo(tmpstr) != 0)
                {
                    tmpcell.SetCellValue("■是");
                }
                else
                {
                    tmprow = sheetSource.GetRow(6);
                    tmpcell = tmprow.GetCell(21);
                    tmpcell.SetCellValue("■否");
                }
                List<IFeature> jzdSorted = pointjzdsorted;
                List<IFeature> jzxSorted = linejzdsorted;

                if (jzxSorted.Count > 9)
                    createRows(sheetSource, 27, jzxSorted.Count - 9);
                int i = 10;
                int tmpInt = -1;
                bool first = true;
                foreach (IFeature f in jzdSorted)
                {
                    if (f != null)
                    {
                        tmprow = sheetSource.GetRow(i);
                        tmpcell = tmprow.GetCell(0);
                        tmpcell.SetCellValue(f.get_Value(jzdhIndex).ToString());

                        tmpstr = f.get_Value(jzdlxIndex).ToString();
                        if (jblxHT[tmpstr] == null)
                            tmpInt = (int)jblxHT["null"];
                        else
                            tmpInt = (int)jblxHT[tmpstr];

                        tmpcell = tmprow.GetCell(tmpInt);
                        tmpcell.SetCellValue("√");
                    }
                    if (first)
                    {
                        i += 1;
                        first = false;
                    }
                    else
                        i += 2;
                }

                i = 10;
                foreach (IFeature f in jzxSorted)
                {
                    if (f != null)
                    {
                        tmprow = sheetSource.GetRow(i);

                        tmpstr = f.get_Value(jzxlbIndex).ToString();
                        if (jzxlxHT[tmpstr] == null)
                            tmpInt = (int)jzxlxHT["null"];
                        else
                            tmpInt = (int)jzxlxHT[tmpstr];

                        tmpcell = tmprow.GetCell(tmpInt);
                        tmpcell.SetCellValue("√");

                        tmpstr = f.get_Value(jzxwzIndex).ToString();
                        if (jzxwzHT[tmpstr] == null)
                            tmpInt = (int)jzxwzHT["null"];
                        else
                            tmpInt = (int)jzxwzHT[tmpstr];

                        tmpcell = tmprow.GetCell(tmpInt);
                        tmpcell.SetCellValue("√");

                        tmpcell = tmprow.GetCell(18);
                        tmpcell.SetCellValue(f.get_Value(jzxsmIndex).ToString());

                        tmpcell = tmprow.GetCell(19);
                        tmpstr = f.get_Value(pldwqlrIndex).ToString();
                        string[] qlrArr = tmpstr.Split(',');
                        //左边0是承包方，那么权利人指界人都是1
                        if (qlrArr[0].CompareTo(cbfmc) == 0)
                        {
                            if (qlrArr[1].CompareTo("") == 0)
                            {
                                tmpcell.SetCellValue("/");
                                tmpcell = tmprow.GetCell(20);
                                tmpcell.SetCellValue(InfoHashTable["fbffzr"] + "");
                            }
                            else
                            {
                                tmpcell.SetCellValue(qlrArr[1]);
                                tmpcell = tmprow.GetCell(20);
                                tmpcell.SetCellValue(qlrArr[1]);
                            }
                        }
                        else//左边0不是承包方，那么权利人指界人都是0
                        {
                            if (qlrArr[0].CompareTo("") == 0)
                            {
                                tmpcell.SetCellValue("/");
                                tmpcell = tmprow.GetCell(20);
                                tmpcell.SetCellValue(InfoHashTable["fbffzr"] + "");
                            }
                            else
                            {
                                tmpcell.SetCellValue(qlrArr[0]);
                                tmpcell = tmprow.GetCell(20);
                                tmpcell.SetCellValue(qlrArr[0]);
                            }
                        }
                    }
                    i += 2;
                }
                i = i > 25 ? i + 1 : 26;
                // sheetSource.AddMergedRegion(new CellRangeAddress(10, i, CommonHelper.Col('V'), CommonHelper.Col('V')));
                //本地块指阶人
                tmprow = sheetSource.GetRow(10);
                tmpcell = tmprow.GetCell(21);
                cbfmc = zdtF.get_Value(cbfmcIndex).ToString();
                tmpcell.SetCellValue(zdtF.get_Value(cbfmcIndex).ToString());
                //调查员
                tmprow = sheetSource.GetRow(31);
                tmpcell = tmprow.GetCell(4);
                tmpcell.SetCellValue(InfoHashTable["dcy"].ToString());
                //审核意见
                tmprow = sheetSource.GetRow(32);
                tmpcell = tmprow.GetCell(1);
                tmpcell.SetCellValue("合格");
                //日期两个
                DateTime cbfdcrq = Convert.ToDateTime(InfoHashTable["dcrq"].ToString());
                TimeSpan timeSpan = new TimeSpan(8, 0, 0, 0);
                DateTime gsshrq = cbfdcrq.Add(timeSpan);
                tmprow = sheetSource.GetRow(31);
                tmpcell = tmprow.GetCell(18);
                tmpcell.SetCellValue(cbfdcrq.ToLongDateString());

                tmprow = sheetSource.GetRow(34);
                tmpcell = tmprow.GetCell(18);
                tmpcell.SetCellValue(gsshrq.ToLongDateString());

                System.IO.FileStream fs = new System.IO.FileStream(outpath, FileMode.Open, FileAccess.ReadWrite);
                workbookSource.Write(fs);
                fs.Close();
                return true;
            }
            return false;
        }
        private bool CreateCM11Table(IFeatureWorkspace workspace, string outpath, IFeature zdtFs, IFeatureClass jzx, IFeatureClass jzd, List<IFeature> pointjzdsorted, List<IFeature> linejzdsorted, Hashtable tempHashTable)
        {
            Hashtable InfoHashTable = tempHashTable;
            IFeature LzdtF = zdtFs;
            IFeature RzdtF = zdtFs;
            Hashtable tdytHashTable = new Hashtable();
            tdytHashTable.Add("1", "种植业");
            tdytHashTable.Add("2", "林业");
            tdytHashTable.Add("3", "畜牧业");
            tdytHashTable.Add("4", "渔业");
            tdytHashTable.Add("5", "其他");

            Hashtable tdlylxHashTable = new Hashtable();
            tdlylxHashTable.Add("011", "水田");
            tdlylxHashTable.Add("012", "水浇地");
            tdlylxHashTable.Add("013", "旱地");

            Hashtable dldjHashTable = new Hashtable();
            dldjHashTable.Add("01", "一等地");
            dldjHashTable.Add("02", "二等地");
            dldjHashTable.Add("03", "三等地");
            dldjHashTable.Add("04", "四等地");
            dldjHashTable.Add("05", "五等地");
            dldjHashTable.Add("06", "六等地");
            dldjHashTable.Add("07", "七等地");

            Hashtable jblxHT = new Hashtable();
            jblxHT.Add("1", 1);
            jblxHT.Add("2", 2);
            jblxHT.Add("3", 3);
            jblxHT.Add("4", 4);
            jblxHT.Add("null", 5);

            Hashtable jzxlxHT = new Hashtable();
            jzxlxHT.Add("01", 6);
            jzxlxHT.Add("02", 7);
            jzxlxHT.Add("03", 8);
            jzxlxHT.Add("04", 9);
            jzxlxHT.Add("05", 10);
            jzxlxHT.Add("06", 11);
            jzxlxHT.Add("07", 12);
            jzxlxHT.Add("08", 13);
            jzxlxHT.Add("09", 14);

            Hashtable jzxwzHT = new Hashtable();
            jzxwzHT.Add("1", 15);
            jzxwzHT.Add("2", 16);
            jzxwzHT.Add("3", 17);



            string templateUrl = AppDomain.CurrentDomain.BaseDirectory + @"template\承包地块调查表.xls";

            //LEFT
            int LfbfbmIndex = LzdtF.Fields.FindField("FBFBM");
            int LcbfmcIndex = LzdtF.Fields.FindField("CBFMC");
            int LdkbhIndex = LzdtF.Fields.FindField("DKBM");
            int LdkmcIndex = LzdtF.Fields.FindField("DKMC");
            int LhtmjIndex = LzdtF.Fields.FindField("HTMJ");
            int LdzIndex = LzdtF.Fields.FindField("DKDZ");
            int LxzIndex = LzdtF.Fields.FindField("DKXZ");
            int LnzIndex = LzdtF.Fields.FindField("DKNZ");
            int LbzIndex = LzdtF.Fields.FindField("DKBZ");
            int LtdytIndex = LzdtF.Fields.FindField("TDYT");
            int LtdlylxIndex = LzdtF.Fields.FindField("TDLYLX");
            int LdldjIndex = LzdtF.Fields.FindField("DLDJ");
            int LsfjbntIndex = LzdtF.Fields.FindField("SFJBNT");
            //RIGHT
            int RfbfbmIndex = RzdtF.Fields.FindField("FBFBM");
            int RcbfmcIndex = RzdtF.Fields.FindField("CBFMC");
            int RdkbhIndex = RzdtF.Fields.FindField("DKBM");
            int RdkmcIndex = RzdtF.Fields.FindField("DKMC");
            int RhtmjIndex = RzdtF.Fields.FindField("HTMJ");
            int RdzIndex = RzdtF.Fields.FindField("DKDZ");
            int RxzIndex = RzdtF.Fields.FindField("DKXZ");
            int RnzIndex = RzdtF.Fields.FindField("DKNZ");
            int RbzIndex = RzdtF.Fields.FindField("DKBZ");
            int RtdytIndex = RzdtF.Fields.FindField("TDYT");
            int RtdlylxIndex = RzdtF.Fields.FindField("TDLYLX");
            int RdldjIndex = RzdtF.Fields.FindField("DLDJ");
            int RsfjbntIndex = RzdtF.Fields.FindField("SFJBNT");



            int jzdlxIndex = jzd.Fields.FindField("JBLX");
            int jzdhIndex = jzd.Fields.FindField("JZDH");
            int jzxlbIndex = jzx.Fields.FindField("JZXLB");
            int jzxwzIndex = jzx.Fields.FindField("JZXWZ");
            int jzxsmIndex = jzx.Fields.FindField("JZXSM");
            int pldwqlrIndex = jzx.Fields.FindField("PLDWQLR");
            int pldwzjrIndex = jzx.Fields.FindField("PLDWZJR");

            // if (File.Exists(outpath)) File.Delete(outpath);
            File.Copy(templateUrl, outpath, true);
            using (System.IO.FileStream fileStream = new System.IO.FileStream(outpath, FileMode.Open, FileAccess.ReadWrite))
            {
                HSSFWorkbook workbookSource = new HSSFWorkbook(fileStream);
                HSSFSheet sheetSource = (HSSFSheet)workbookSource.GetSheetAt(0);
                //设定合并单元格的样式
                HSSFCellStyle style = (HSSFCellStyle)workbookSource.CreateCellStyle();
                style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.CENTER;
                style.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.CENTER;
                style.BorderBottom = NPOI.SS.UserModel.BorderStyle.THIN;
                style.BorderRight = NPOI.SS.UserModel.BorderStyle.THIN;
                style.BorderLeft = NPOI.SS.UserModel.BorderStyle.THIN;
                style.BorderTop = NPOI.SS.UserModel.BorderStyle.THIN;
                style.WrapText = true;
                string cbfmc = null;

                var tmprow = sheetSource.GetRow(2);
                //L发包方编码                
                var tmpcell = tmprow.GetCell(4);
                tmpcell.SetCellValue(LzdtF.get_Value(LfbfbmIndex).ToString());
                //R发包方编码
                tmpcell = tmprow.GetCell(27);
                tmpcell.SetCellValue(RzdtF.get_Value(RfbfbmIndex).ToString());


                //L承包方代表
                tmpcell = tmprow.GetCell(12);
                cbfmc = LzdtF.get_Value(LcbfmcIndex).ToString();
                tmpcell.SetCellValue(LzdtF.get_Value(LcbfmcIndex).ToString());
                //R承包方代表
                tmpcell = tmprow.GetCell(12 + 23);
                cbfmc = RzdtF.get_Value(RcbfmcIndex).ToString();
                tmpcell.SetCellValue(RzdtF.get_Value(RcbfmcIndex).ToString());
                //+++++++++++++图幅编号在这里加++++++++++++++
                tmprow = sheetSource.GetRow(3);
                //地块编号
                tmpcell = tmprow.GetCell(4);
                tmpcell.SetCellValue(LzdtF.get_Value(LdkbhIndex).ToString().Substring(14));
                tmpcell = tmprow.GetCell(27);
                tmpcell.SetCellValue(RzdtF.get_Value(RdkbhIndex).ToString().Substring(14));
                //地块名称
                tmpcell = tmprow.GetCell(12);
                tmpcell.SetCellValue(LzdtF.get_Value(LdkmcIndex).ToString());
                tmpcell = tmprow.GetCell(35);
                tmpcell.SetCellValue(RzdtF.get_Value(RdkmcIndex).ToString());
                //合同面积
                tmpcell = tmprow.GetCell(20);
                tmpcell.SetCellValue(((double)LzdtF.get_Value(LhtmjIndex)).ToString("f"));
                tmpcell = tmprow.GetCell(43);
                tmpcell.SetCellValue(((double)RzdtF.get_Value(RhtmjIndex)).ToString("f"));

                tmprow = sheetSource.GetRow(4);
                //东至
                tmpcell = tmprow.GetCell(1);
                tmpcell.SetCellValue(LzdtF.get_Value(LdzIndex).ToString());
                tmpcell = tmprow.GetCell(24);
                tmpcell.SetCellValue(RzdtF.get_Value(RdzIndex).ToString());
                //南至
                tmpcell = tmprow.GetCell(9);
                tmpcell.SetCellValue(LzdtF.get_Value(LnzIndex).ToString());
                tmpcell = tmprow.GetCell(32);
                tmpcell.SetCellValue(RzdtF.get_Value(RnzIndex).ToString());
                //西至
                tmpcell = tmprow.GetCell(16);
                tmpcell.SetCellValue(LzdtF.get_Value(LxzIndex).ToString());
                tmpcell = tmprow.GetCell(39);
                tmpcell.SetCellValue(RzdtF.get_Value(RxzIndex).ToString());
                //北至
                tmpcell = tmprow.GetCell(20);
                tmpcell.SetCellValue(LzdtF.get_Value(LbzIndex).ToString());
                tmpcell = tmprow.GetCell(43);
                tmpcell.SetCellValue(RzdtF.get_Value(RbzIndex).ToString());

                tmprow = sheetSource.GetRow(5);
                //土地用途
                tmpcell = tmprow.GetCell(1);
                string tmpstr = tdytHashTable[LzdtF.get_Value(LtdytIndex)].ToString();
                IRichTextString rich = tmpcell.RichStringCellValue;
                IFont font = workbookSource.GetFontAt(rich.GetFontAtIndex(rich.String.Length - 1));
                tmpstr = rich.String.Replace("□" + tmpstr, "■" + tmpstr);
                rich = new HSSFRichTextString(tmpstr);
                rich.ApplyFont(tmpstr.IndexOf('_'), tmpstr.LastIndexOf('_'), font);
                tmpcell.SetCellValue(rich);

                tmpcell = tmprow.GetCell(24);
                tmpstr = tdytHashTable[RzdtF.get_Value(RtdytIndex)].ToString();
                rich = tmpcell.RichStringCellValue;
                font = workbookSource.GetFontAt(rich.GetFontAtIndex(rich.String.Length - 1));
                tmpstr = rich.String.Replace("□" + tmpstr, "■" + tmpstr);
                rich = new HSSFRichTextString(tmpstr);
                rich.ApplyFont(tmpstr.IndexOf('_'), tmpstr.LastIndexOf('_'), font);
                tmpcell.SetCellValue(rich);
                //土地利用类型
                tmpcell = tmprow.GetCell(15);
                tmpstr = tdlylxHashTable[LzdtF.get_Value(LtdlylxIndex)].ToString();
                rich = tmpcell.RichStringCellValue;
                font = workbookSource.GetFontAt(rich.GetFontAtIndex(rich.String.Length - 1));
                tmpstr = rich.String.Replace("□" + tmpstr, "■" + tmpstr);
                rich = new HSSFRichTextString(tmpstr);
                rich.ApplyFont(tmpstr.IndexOf('_'), tmpstr.LastIndexOf('_'), font);
                tmpcell.SetCellValue(rich);

                tmpcell = tmprow.GetCell(38);
                tmpstr = tdlylxHashTable[RzdtF.get_Value(RtdlylxIndex)].ToString();
                rich = tmpcell.RichStringCellValue;
                font = workbookSource.GetFontAt(rich.GetFontAtIndex(rich.String.Length - 1));
                tmpstr = rich.String.Replace("□" + tmpstr, "■" + tmpstr);
                rich = new HSSFRichTextString(tmpstr);
                rich.ApplyFont(tmpstr.IndexOf('_'), tmpstr.LastIndexOf('_'), font);
                tmpcell.SetCellValue(rich);
                //利用类型
                tmpcell = tmprow.GetCell(10);
                tmpstr = dldjHashTable[LzdtF.get_Value(LdldjIndex)].ToString();
                tmpcell.SetCellValue(tmpstr);

                tmpcell = tmprow.GetCell(33);
                tmpstr = dldjHashTable[RzdtF.get_Value(RdldjIndex)].ToString();
                tmpcell.SetCellValue(tmpstr);
                //是否基本农田
                tmpcell = tmprow.GetCell(21);
                tmpstr = LzdtF.get_Value(LsfjbntIndex).ToString();
                if ("2".CompareTo(tmpstr) != 0)
                {
                    tmpcell.SetCellValue("■是");
                }
                else
                {
                    tmprow = sheetSource.GetRow(6);
                    tmpcell = tmprow.GetCell(21);
                    tmpcell.SetCellValue("■否");
                }

                tmpcell = tmprow.GetCell(44);
                tmpstr = RzdtF.get_Value(RsfjbntIndex).ToString();
                if ("2".CompareTo(tmpstr) != 0)
                {
                    tmpcell.SetCellValue("■是");
                }
                else
                {
                    tmprow = sheetSource.GetRow(6);
                    tmpcell = tmprow.GetCell(21);
                    tmpcell.SetCellValue("■否");
                }
                //////界址点线////////////////////////////////////////////////////////
                //拿出地块界址点线
                List<IFeature> LjzdSorted = new List<IFeature>();
                List<IFeature> LjzxSorted = new List<IFeature>();
                List<IFeature> RjzdSorted = new List<IFeature>();
                List<IFeature> RjzxSorted = new List<IFeature>();
                for (int i = 0; i < 9; i++)
                {
                    LjzdSorted.Add(pointjzdsorted[i]);
                    LjzxSorted.Add(linejzdsorted[i]);
                }
                for (int i = 9; i < pointjzdsorted.Count; i++)
                {
                    RjzdSorted.Add(pointjzdsorted[i]);
                    RjzxSorted.Add(linejzdsorted[i]);
                }
                //如果界址点数量大于九需要加方格数量错误！应该启下一页
                if (LjzxSorted.Count > 9)
                    createRows(sheetSource, 27, LjzxSorted.Count - 9);
                ///.......................................................
                int Li = 10;
                int LtmpInt = -1;
                bool Lfirst = true;
                foreach (IFeature Lf in LjzdSorted)
                {
                    if (Lf != null)
                    {
                        tmprow = sheetSource.GetRow(Li);
                        tmpcell = tmprow.GetCell(0);
                        tmpcell.SetCellValue(Lf.get_Value(jzdhIndex).ToString());

                        tmpstr = Lf.get_Value(jzdlxIndex).ToString();
                        if (jblxHT[tmpstr] == null)
                            LtmpInt = (int)jblxHT["null"];
                        else
                            LtmpInt = (int)jblxHT[tmpstr];

                        tmpcell = tmprow.GetCell(LtmpInt);
                        tmpcell.SetCellValue("√");
                    }
                    if (Lfirst)
                    {
                        Li += 1;
                        Lfirst = false;
                    }
                    else
                        Li += 2;
                }

                Li = 10;
                foreach (IFeature Lf in LjzxSorted)
                {
                    if (Lf != null)
                    {
                        tmprow = sheetSource.GetRow(Li);

                        tmpstr = Lf.get_Value(jzxlbIndex).ToString();
                        if (jzxlxHT[tmpstr] == null)
                            LtmpInt = (int)jzxlxHT["null"];
                        else
                            LtmpInt = (int)jzxlxHT[tmpstr];

                        tmpcell = tmprow.GetCell(LtmpInt);
                        tmpcell.SetCellValue("√");

                        tmpstr = Lf.get_Value(jzxwzIndex).ToString();
                        if (jzxwzHT[tmpstr] == null)
                            LtmpInt = (int)jzxwzHT["null"];
                        else
                            LtmpInt = (int)jzxwzHT[tmpstr];

                        tmpcell = tmprow.GetCell(LtmpInt);
                        tmpcell.SetCellValue("√");

                        tmpcell = tmprow.GetCell(18);
                        tmpcell.SetCellValue(Lf.get_Value(jzxsmIndex).ToString());

                        tmpcell = tmprow.GetCell(19);
                        tmpstr = Lf.get_Value(pldwqlrIndex).ToString();
                        string[] qlrArr = tmpstr.Split(',');
                        //左边0是承包方，那么权利人指界人都是1
                        if (qlrArr[0].CompareTo(cbfmc) == 0)
                        {
                            if (qlrArr[1].CompareTo("") == 0)
                            {
                                tmpcell.SetCellValue("/");
                                tmpcell = tmprow.GetCell(20);
                                tmpcell.SetCellValue(InfoHashTable["fbffzr"] + "");
                            }
                            else
                            {
                                tmpcell.SetCellValue(qlrArr[1]);
                                tmpcell = tmprow.GetCell(20);
                                tmpcell.SetCellValue(qlrArr[1]);
                            }
                        }
                        else//左边0不是承包方，那么权利人指界人都是0
                        {
                            if (qlrArr[0].CompareTo("") == 0)
                            {
                                tmpcell.SetCellValue("/");
                                tmpcell = tmprow.GetCell(20);
                                tmpcell.SetCellValue(InfoHashTable["fbffzr"] + "");
                            }
                            else
                            {
                                tmpcell.SetCellValue(qlrArr[0]);
                                tmpcell = tmprow.GetCell(20);
                                tmpcell.SetCellValue(qlrArr[0]);
                            }
                        }
                    }
                    Li += 2;
                }
                Li = Li > 25 ? Li + 1 : 26;
                //sheetSource.AddMergedRegion(new CellRangeAddress(10, Li, CommonHelper.Col('V'), CommonHelper.Col('V')));

                /////RRRRRRRRRRRRRRRRRRRRRRR
                if (RjzxSorted.Count > 9)
                    createRows(sheetSource, 27, RjzxSorted.Count - 9);
                int Ri = 10;
                int RtmpInt = -1;
                bool Rfirst = true;
                foreach (IFeature Rf in RjzdSorted)
                {
                    if (Rf != null)
                    {
                        tmprow = sheetSource.GetRow(Ri);
                        tmpcell = tmprow.GetCell(23);
                        tmpcell.SetCellValue(Rf.get_Value(jzdhIndex).ToString());

                        tmpstr = Rf.get_Value(jzdlxIndex).ToString();
                        if (jblxHT[tmpstr] == null)
                            RtmpInt = (int)jblxHT["null"];
                        else
                            RtmpInt = (int)jblxHT[tmpstr];

                        tmpcell = tmprow.GetCell(RtmpInt + 23);
                        tmpcell.SetCellValue("√");
                    }
                    if (Rfirst)
                    {
                        Ri += 1;
                        Rfirst = false;
                    }
                    else
                        Ri += 2;
                }

                Ri = 10;
                foreach (IFeature Rf in RjzxSorted)
                {
                    if (Rf != null)
                    {
                        tmprow = sheetSource.GetRow(Ri);

                        tmpstr = Rf.get_Value(jzxlbIndex).ToString();
                        if (jzxlxHT[tmpstr] == null)
                            RtmpInt = (int)jzxlxHT["null"];
                        else
                            RtmpInt = (int)jzxlxHT[tmpstr];

                        tmpcell = tmprow.GetCell(RtmpInt + 23);
                        tmpcell.SetCellValue("√");

                        tmpstr = Rf.get_Value(jzxwzIndex).ToString();
                        if (jzxwzHT[tmpstr] == null)
                            RtmpInt = (int)jzxwzHT["null"];
                        else
                            RtmpInt = (int)jzxwzHT[tmpstr];

                        tmpcell = tmprow.GetCell(RtmpInt + 23);
                        tmpcell.SetCellValue("√");

                        tmpcell = tmprow.GetCell(18 + 23);
                        tmpcell.SetCellValue(Rf.get_Value(jzxsmIndex).ToString());

                        tmpcell = tmprow.GetCell(19 + 23);
                        tmpstr = Rf.get_Value(pldwqlrIndex).ToString();
                        string[] qlrArr = tmpstr.Split(',');
                        //左边0是承包方，那么权利人指界人都是1
                        if (qlrArr[0].CompareTo(cbfmc) == 0)
                        {
                            if (qlrArr[1].CompareTo("") == 0)
                            {
                                tmpcell.SetCellValue("/");
                                tmpcell = tmprow.GetCell(20 + 23);
                                tmpcell.SetCellValue(InfoHashTable["fbffzr"] + "");
                            }
                            else
                            {
                                tmpcell.SetCellValue(qlrArr[1]);
                                tmpcell = tmprow.GetCell(20 + 23);
                                tmpcell.SetCellValue(qlrArr[1]);
                            }
                        }
                        else//左边0不是承包方，那么权利人指界人都是0
                        {
                            if (qlrArr[0].CompareTo("") == 0)
                            {
                                tmpcell.SetCellValue("/");
                                tmpcell = tmprow.GetCell(20 + 23);
                                tmpcell.SetCellValue(InfoHashTable["fbffzr"] + "");
                            }
                            else
                            {
                                tmpcell.SetCellValue(qlrArr[0]);
                                tmpcell = tmprow.GetCell(20 + 23);
                                tmpcell.SetCellValue(qlrArr[0]);
                            }
                        }
                    }
                    Ri += 2;
                }
                Ri = Ri > 25 ? Ri + 1 : 26;
                //sheetSource.AddMergedRegion(new CellRangeAddress(10, Ri, CommonHelper.Col('V'), CommonHelper.Col('V')));
                //本地块指界人
                tmprow = sheetSource.GetRow(10);
                tmpcell = tmprow.GetCell(21);
                cbfmc = LzdtF.get_Value(LcbfmcIndex).ToString();
                tmpcell.SetCellValue(LzdtF.get_Value(LcbfmcIndex).ToString());

                tmpcell = tmprow.GetCell(21 + 23);
                cbfmc = RzdtF.get_Value(RcbfmcIndex).ToString();
                tmpcell.SetCellValue(RzdtF.get_Value(RcbfmcIndex).ToString());
                //lllllllllllllllllllllllll调查员
                tmprow = sheetSource.GetRow(31);
                tmpcell = tmprow.GetCell(4);
                tmpcell.SetCellValue(InfoHashTable["dcy"].ToString());
                //审核意见
                tmprow = sheetSource.GetRow(32);
                tmpcell = tmprow.GetCell(1);
                tmpcell.SetCellValue("合格");
                //日期两个
                //日期两个
                DateTime cbfdcrq = Convert.ToDateTime(InfoHashTable["dcrq"].ToString());
                TimeSpan timeSpan = new TimeSpan(8, 0, 0, 0);
                DateTime gsshrq = cbfdcrq.Add(timeSpan);
                tmprow = sheetSource.GetRow(31);
                tmpcell = tmprow.GetCell(18);
                tmpcell.SetCellValue(cbfdcrq.ToLongDateString());

                tmprow = sheetSource.GetRow(34);
                tmpcell = tmprow.GetCell(18);
                tmpcell.SetCellValue(gsshrq.ToLongDateString());

                //rrrrrrrrrrrrrrrrrrrrrrrr调查员
                tmprow = sheetSource.GetRow(31);
                tmpcell = tmprow.GetCell(4 + 23);
                tmpcell.SetCellValue(InfoHashTable["dcy"].ToString());
                //审核意见
                tmprow = sheetSource.GetRow(32);
                tmpcell = tmprow.GetCell(1 + 23);
                tmpcell.SetCellValue("合格");

                tmprow = sheetSource.GetRow(31);
                tmpcell = tmprow.GetCell(18 + 23);
                tmpcell.SetCellValue(cbfdcrq.ToLongDateString());

                tmprow = sheetSource.GetRow(34);
                tmpcell = tmprow.GetCell(18 + 23);
                tmpcell.SetCellValue(gsshrq.ToLongDateString());

                //保存
                System.IO.FileStream fs = new System.IO.FileStream(outpath, FileMode.Open, FileAccess.ReadWrite);
                workbookSource.Write(fs);
                fs.Close();
                return true;
            }
            return false;
        }
        private IFeature getRelationLine(IGeometry query, List<IFeature> fList)
        {
            IRelationalOperator rel = query as IRelationalOperator;
            IGeometry tmpgeo = null;
            foreach (IFeature f in fList)
            {
                tmpgeo = f.ShapeCopy;
                if (rel.Touches(tmpgeo))
                {
                    return f;
                    break;
                }
            }
            return null;
        }

        private IFeature getRelationPoint(IGeometry query, List<IFeature> fList)
        {
            IRelationalOperator rel = query as IRelationalOperator;
            foreach (IFeature f in fList)
            {
                if (rel.Touches(f.ShapeCopy))
                {
                    return f;
                    break;
                }
            }
            return null;
        }

        private bool CreateOneC2Table(string databaseUrl, int[] fids, string outpath, string zdtname, string jzxname, string jzdname, Hashtable InfoHashTable)
        {
            try
            {
                Hashtable tempHashTable = InfoHashTable;
                IAeFactory pAeFactory = new PersonalGeoDatabase(databaseUrl);
                IFeatureWorkspace workspace = pAeFactory.OpenFeatrueWorkspace();
                IFeatureClass zdt = workspace.OpenFeatureClass(zdtname);
                IFeatureClass jzx = workspace.OpenFeatureClass(jzxname);
                IFeatureClass jzd = workspace.OpenFeatureClass(jzdname);
                List<IFeature> feas = new List<IFeature>();
                feas.Add(zdt.GetFeature(fids[0]));
                feas.Add(zdt.GetFeature(fids[1]));
                var ret = CreateC2Table(workspace, outpath, feas, jzx, jzd, tempHashTable);
                Marshal.FinalReleaseComObject(zdt);
                Marshal.FinalReleaseComObject(workspace);
                GC.WaitForPendingFinalizers();
                GC.Collect();
                return ret;
            }
            catch
            {
            }
            return false;
        }

        private bool CreateC2Table(IFeatureWorkspace workspace, string outpath, List<IFeature> zdtFs, IFeatureClass jzx, IFeatureClass jzd, Hashtable InfoHashTable)
        {
            IFeature LzdtF = zdtFs[0];
            IFeature RzdtF = zdtFs[1];
            Hashtable tdytHashTable = new Hashtable();
            tdytHashTable.Add("1", "种植业");
            tdytHashTable.Add("2", "林业");
            tdytHashTable.Add("3", "畜牧业");
            tdytHashTable.Add("4", "渔业");
            tdytHashTable.Add("5", "其他");

            Hashtable tdlylxHashTable = new Hashtable();
            tdlylxHashTable.Add("011", "水田");
            tdlylxHashTable.Add("012", "水浇地");
            tdlylxHashTable.Add("013", "旱地");

            Hashtable dldjHashTable = new Hashtable();
            dldjHashTable.Add("01", "一等地");
            dldjHashTable.Add("02", "二等地");
            dldjHashTable.Add("03", "三等地");
            dldjHashTable.Add("04", "四等地");
            dldjHashTable.Add("05", "五等地");
            dldjHashTable.Add("06", "六等地");
            dldjHashTable.Add("07", "七等地");

            Hashtable jblxHT = new Hashtable();
            jblxHT.Add("1", 1);
            jblxHT.Add("2", 2);
            jblxHT.Add("3", 3);
            jblxHT.Add("4", 4);
            jblxHT.Add("null", 5);

            Hashtable jzxlxHT = new Hashtable();
            jzxlxHT.Add("01", 6);
            jzxlxHT.Add("02", 7);
            jzxlxHT.Add("03", 8);
            jzxlxHT.Add("04", 9);
            jzxlxHT.Add("05", 10);
            jzxlxHT.Add("06", 11);
            jzxlxHT.Add("07", 12);
            jzxlxHT.Add("08", 13);
            jzxlxHT.Add("09", 14);

            Hashtable jzxwzHT = new Hashtable();
            jzxwzHT.Add("1", 15);
            jzxwzHT.Add("2", 16);
            jzxwzHT.Add("3", 17);



            string templateUrl = AppDomain.CurrentDomain.BaseDirectory + @"template\承包地块调查表.xls";

            //LEFT
            int LfbfbmIndex = LzdtF.Fields.FindField("FBFBM");
            int LcbfmcIndex = LzdtF.Fields.FindField("CBFMC");
            int LdkbhIndex = LzdtF.Fields.FindField("DKBM");
            int LdkmcIndex = LzdtF.Fields.FindField("DKMC");
            int LhtmjIndex = LzdtF.Fields.FindField("HTMJ");
            int LdzIndex = LzdtF.Fields.FindField("DKDZ");
            int LxzIndex = LzdtF.Fields.FindField("DKXZ");
            int LnzIndex = LzdtF.Fields.FindField("DKNZ");
            int LbzIndex = LzdtF.Fields.FindField("DKBZ");
            int LtdytIndex = LzdtF.Fields.FindField("TDYT");
            int LtdlylxIndex = LzdtF.Fields.FindField("TDLYLX");
            int LdldjIndex = LzdtF.Fields.FindField("DLDJ");
            int LsfjbntIndex = LzdtF.Fields.FindField("SFJBNT");
            //RIGHT
            int RfbfbmIndex = RzdtF.Fields.FindField("FBFBM");
            int RcbfmcIndex = RzdtF.Fields.FindField("CBFMC");
            int RdkbhIndex = RzdtF.Fields.FindField("DKBM");
            int RdkmcIndex = RzdtF.Fields.FindField("DKMC");
            int RhtmjIndex = RzdtF.Fields.FindField("HTMJ");
            int RdzIndex = RzdtF.Fields.FindField("DKDZ");
            int RxzIndex = RzdtF.Fields.FindField("DKXZ");
            int RnzIndex = RzdtF.Fields.FindField("DKNZ");
            int RbzIndex = RzdtF.Fields.FindField("DKBZ");
            int RtdytIndex = RzdtF.Fields.FindField("TDYT");
            int RtdlylxIndex = RzdtF.Fields.FindField("TDLYLX");
            int RdldjIndex = RzdtF.Fields.FindField("DLDJ");
            int RsfjbntIndex = RzdtF.Fields.FindField("SFJBNT");



            int jzdlxIndex = jzd.Fields.FindField("JBLX");
            int jzdhIndex = jzd.Fields.FindField("JZDH");
            int jzxlbIndex = jzx.Fields.FindField("JZXLB");
            int jzxwzIndex = jzx.Fields.FindField("JZXWZ");
            int jzxsmIndex = jzx.Fields.FindField("JZXSM");
            int pldwqlrIndex = jzx.Fields.FindField("PLDWQLR");
            int pldwzjrIndex = jzx.Fields.FindField("PLDWZJR");

            // if (File.Exists(outpath)) File.Delete(outpath);
            File.Copy(templateUrl, outpath, true);
            using (System.IO.FileStream fileStream = new System.IO.FileStream(outpath, FileMode.Open, FileAccess.ReadWrite))
            {
                HSSFWorkbook workbookSource = new HSSFWorkbook(fileStream);
                HSSFSheet sheetSource = (HSSFSheet)workbookSource.GetSheetAt(0);
                //设定合并单元格的样式
                HSSFCellStyle style = (HSSFCellStyle)workbookSource.CreateCellStyle();
                style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.CENTER;
                style.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.CENTER;
                style.BorderBottom = NPOI.SS.UserModel.BorderStyle.THIN;
                style.BorderRight = NPOI.SS.UserModel.BorderStyle.THIN;
                style.BorderLeft = NPOI.SS.UserModel.BorderStyle.THIN;
                style.BorderTop = NPOI.SS.UserModel.BorderStyle.THIN;
                style.WrapText = true;
                string cbfmc = null;

                var tmprow = sheetSource.GetRow(2);
                //L发包方编码                
                var tmpcell = tmprow.GetCell(4);
                tmpcell.SetCellValue(LzdtF.get_Value(LfbfbmIndex).ToString());
                //R发包方编码
                tmpcell = tmprow.GetCell(27);
                tmpcell.SetCellValue(RzdtF.get_Value(RfbfbmIndex).ToString());


                //L承包方代表
                tmpcell = tmprow.GetCell(12);
                cbfmc = LzdtF.get_Value(LcbfmcIndex).ToString();
                tmpcell.SetCellValue(LzdtF.get_Value(LcbfmcIndex).ToString());
                //R承包方代表
                tmpcell = tmprow.GetCell(12 + 23);
                cbfmc = RzdtF.get_Value(RcbfmcIndex).ToString();
                tmpcell.SetCellValue(RzdtF.get_Value(RcbfmcIndex).ToString());
                //+++++++++++++图幅编号在这里加++++++++++++++
                tmprow = sheetSource.GetRow(3);
                //地块编号
                tmpcell = tmprow.GetCell(4);
                tmpcell.SetCellValue(LzdtF.get_Value(LdkbhIndex).ToString().Substring(14));
                tmpcell = tmprow.GetCell(27);
                tmpcell.SetCellValue(RzdtF.get_Value(RdkbhIndex).ToString().Substring(14));
                //地块名称
                tmpcell = tmprow.GetCell(12);
                tmpcell.SetCellValue(LzdtF.get_Value(LdkmcIndex).ToString());
                tmpcell = tmprow.GetCell(35);
                tmpcell.SetCellValue(RzdtF.get_Value(RdkmcIndex).ToString());
                //合同面积
                tmpcell = tmprow.GetCell(20);
                tmpcell.SetCellValue(((double)LzdtF.get_Value(LhtmjIndex)).ToString("f"));
                tmpcell = tmprow.GetCell(43);
                tmpcell.SetCellValue(((double)RzdtF.get_Value(RhtmjIndex)).ToString("f"));

                tmprow = sheetSource.GetRow(4);
                //东至
                tmpcell = tmprow.GetCell(1);
                tmpcell.SetCellValue(LzdtF.get_Value(LdzIndex).ToString());
                tmpcell = tmprow.GetCell(24);
                tmpcell.SetCellValue(RzdtF.get_Value(RdzIndex).ToString());
                //南至
                tmpcell = tmprow.GetCell(9);
                tmpcell.SetCellValue(LzdtF.get_Value(LnzIndex).ToString());
                tmpcell = tmprow.GetCell(32);
                tmpcell.SetCellValue(RzdtF.get_Value(RnzIndex).ToString());
                //西至
                tmpcell = tmprow.GetCell(16);
                tmpcell.SetCellValue(LzdtF.get_Value(LxzIndex).ToString());
                tmpcell = tmprow.GetCell(39);
                tmpcell.SetCellValue(RzdtF.get_Value(RxzIndex).ToString());
                //北至
                tmpcell = tmprow.GetCell(20);
                tmpcell.SetCellValue(LzdtF.get_Value(LbzIndex).ToString());
                tmpcell = tmprow.GetCell(43);
                tmpcell.SetCellValue(RzdtF.get_Value(RbzIndex).ToString());

                tmprow = sheetSource.GetRow(5);
                //土地用途
                tmpcell = tmprow.GetCell(1);
                string tmpstr = tdytHashTable[LzdtF.get_Value(LtdytIndex)].ToString();
                IRichTextString rich = tmpcell.RichStringCellValue;
                IFont font = workbookSource.GetFontAt(rich.GetFontAtIndex(rich.String.Length - 1));
                tmpstr = rich.String.Replace("□" + tmpstr, "■" + tmpstr);
                rich = new HSSFRichTextString(tmpstr);
                rich.ApplyFont(tmpstr.IndexOf('_'), tmpstr.LastIndexOf('_'), font);
                tmpcell.SetCellValue(rich);

                tmpcell = tmprow.GetCell(24);
                tmpstr = tdytHashTable[RzdtF.get_Value(RtdytIndex)].ToString();
                rich = tmpcell.RichStringCellValue;
                font = workbookSource.GetFontAt(rich.GetFontAtIndex(rich.String.Length - 1));
                tmpstr = rich.String.Replace("□" + tmpstr, "■" + tmpstr);
                rich = new HSSFRichTextString(tmpstr);
                rich.ApplyFont(tmpstr.IndexOf('_'), tmpstr.LastIndexOf('_'), font);
                tmpcell.SetCellValue(rich);
                //土地利用类型
                tmpcell = tmprow.GetCell(15);
                tmpstr = tdlylxHashTable[LzdtF.get_Value(LtdlylxIndex)].ToString();
                rich = tmpcell.RichStringCellValue;
                font = workbookSource.GetFontAt(rich.GetFontAtIndex(rich.String.Length - 1));
                tmpstr = rich.String.Replace("□" + tmpstr, "■" + tmpstr);
                rich = new HSSFRichTextString(tmpstr);
                rich.ApplyFont(tmpstr.IndexOf('_'), tmpstr.LastIndexOf('_'), font);
                tmpcell.SetCellValue(rich);

                tmpcell = tmprow.GetCell(38);
                tmpstr = tdlylxHashTable[RzdtF.get_Value(RtdlylxIndex)].ToString();
                rich = tmpcell.RichStringCellValue;
                font = workbookSource.GetFontAt(rich.GetFontAtIndex(rich.String.Length - 1));
                tmpstr = rich.String.Replace("□" + tmpstr, "■" + tmpstr);
                rich = new HSSFRichTextString(tmpstr);
                rich.ApplyFont(tmpstr.IndexOf('_'), tmpstr.LastIndexOf('_'), font);
                tmpcell.SetCellValue(rich);
                //利用类型
                tmpcell = tmprow.GetCell(10);
                tmpstr = dldjHashTable[LzdtF.get_Value(LdldjIndex)].ToString();
                tmpcell.SetCellValue(tmpstr);

                tmpcell = tmprow.GetCell(33);
                tmpstr = dldjHashTable[RzdtF.get_Value(RdldjIndex)].ToString();
                tmpcell.SetCellValue(tmpstr);
                //是否基本农田
                tmpcell = tmprow.GetCell(21);
                tmpstr = LzdtF.get_Value(LsfjbntIndex).ToString();
                if ("2".CompareTo(tmpstr) != 0)
                {
                    tmpcell.SetCellValue("■是");
                }
                else
                {
                    tmprow = sheetSource.GetRow(6);
                    tmpcell = tmprow.GetCell(21);
                    tmpcell.SetCellValue("■否");
                }

                tmpcell = tmprow.GetCell(44);
                tmpstr = RzdtF.get_Value(RsfjbntIndex).ToString();
                if ("2".CompareTo(tmpstr) != 0)
                {
                    tmpcell.SetCellValue("■是");
                }
                else
                {
                    tmprow = sheetSource.GetRow(6);
                    tmpcell = tmprow.GetCell(21);
                    tmpcell.SetCellValue("■否");
                }
                //////界址点线////////////////////////////////////////////////////////
                //拿出地块界址点线
                ISpatialFilter Lsf = new SpatialFilterClass();
                Lsf.Geometry = LzdtF.ShapeCopy;
                Lsf.GeometryField = "SHAPE";
                Lsf.SpatialRel = esriSpatialRelEnum.esriSpatialRelRelation;
                Lsf.SpatialRelDescription = "F*TT*TF*T";
                var Lcursor = jzd.Search(Lsf, false);
                IFeature LtmpFeature = null;
                List<IFeature> LjzdList = new List<IFeature>();
                while ((LtmpFeature = Lcursor.NextFeature()) != null)
                {
                    LjzdList.Add(LtmpFeature);
                }
                IGeometryCollection Lgc = LzdtF.ShapeCopy as IGeometryCollection;
                if (Lgc.GeometryCount > 1)
                {
                    int a = Lgc.GeometryCount;
                }
                Lsf = new SpatialFilterClass();
                Lsf.GeometryField = "SHAPE";
                Lsf.Geometry = LzdtF.ShapeCopy;
                Lsf.SpatialRel = esriSpatialRelEnum.esriSpatialRelRelation;
                Lsf.SpatialRelDescription = "FFTTT*FF*";

                Lcursor = jzx.Search(Lsf, false);
                LtmpFeature = null;
                List<IFeature> LjzxList = new List<IFeature>();

                while ((LtmpFeature = Lcursor.NextFeature()) != null)
                {
                    LjzxList.Add(LtmpFeature);
                }

                List<IFeature> LjzdSorted = new List<IFeature>();
                List<IFeature> LjzxSorted = new List<IFeature>();

                if (LjzdList.Count != LjzxList.Count)
                {
                    // System.Windows.Forms.MessageBox.Show("筛选出的界址点与界址线数量不同！");
                    return false;
                }

                int Lj = LjzdList.Count;
                IFeature Ltmppoint = LjzdList[0];
                IFeature Ltmpline = null;
                int LgcCount = (LzdtF.ShapeCopy as IGeometryCollection).GeometryCount - 1;
                while (LjzxList.Count > 0)
                {
                    if (Ltmppoint == null)
                    {
                        if (LgcCount > 0)
                        {
                            LjzxSorted.Add(null);
                            LjzdSorted.Add(null);
                            Ltmppoint = LjzdList[0];
                            --LgcCount;
                        }
                        else
                        {
                            //  System.Windows.Forms.MessageBox.Show("筛选出的界址点与界址线错误！");
                            return false;
                        }
                    }

                    Ltmpline = getRelationLine(Ltmppoint.ShapeCopy, LjzxList);
                    if (Ltmpline == null)
                    {

                        //  System.Windows.Forms.MessageBox.Show("筛选出的界址点与界址线错误！");
                        return false;
                    }

                    LjzdSorted.Add(Ltmppoint);
                    LjzdList.Remove(Ltmppoint);
                    LjzxSorted.Add(Ltmpline);
                    LjzxList.Remove(Ltmpline);

                    Ltmppoint = getRelationPoint(Ltmpline.ShapeCopy, LjzdList);

                }
                //Right界址点线
                ISpatialFilter Rsf = new SpatialFilterClass();
                Rsf.Geometry = RzdtF.ShapeCopy;
                Rsf.GeometryField = "SHAPE";
                Rsf.SpatialRel = esriSpatialRelEnum.esriSpatialRelRelation;
                Rsf.SpatialRelDescription = "F*TT*TF*T";
                var Rcursor = jzd.Search(Rsf, false);
                IFeature RtmpFeature = null;
                List<IFeature> RjzdList = new List<IFeature>();
                while ((RtmpFeature = Rcursor.NextFeature()) != null)
                {
                    RjzdList.Add(RtmpFeature);
                }
                IGeometryCollection Rgc = RzdtF.ShapeCopy as IGeometryCollection;
                if (Rgc.GeometryCount > 1)
                {
                    int Ra = Rgc.GeometryCount;
                }
                Rsf = new SpatialFilterClass();
                Rsf.GeometryField = "SHAPE";
                Rsf.Geometry = RzdtF.ShapeCopy;
                Rsf.SpatialRel = esriSpatialRelEnum.esriSpatialRelRelation;
                Rsf.SpatialRelDescription = "FFTTT*FF*";

                Rcursor = jzx.Search(Rsf, false);
                RtmpFeature = null;
                List<IFeature> RjzxList = new List<IFeature>();

                while ((RtmpFeature = Rcursor.NextFeature()) != null)
                {
                    RjzxList.Add(RtmpFeature);
                }

                List<IFeature> RjzdSorted = new List<IFeature>();
                List<IFeature> RjzxSorted = new List<IFeature>();

                if (RjzdList.Count != RjzxList.Count)
                {
                    // System.Windows.Forms.MessageBox.Show("筛选出的界址点与界址线数量不同！");
                    return false;
                }

                int Rj = RjzdList.Count;
                IFeature Rtmppoint = RjzdList[0];
                IFeature Rtmpline = null;
                int RgcCount = (RzdtF.ShapeCopy as IGeometryCollection).GeometryCount - 1;
                while (RjzxList.Count > 0)
                {
                    if (Rtmppoint == null)
                    {
                        if (RgcCount > 0)
                        {
                            RjzxSorted.Add(null);
                            RjzdSorted.Add(null);
                            Rtmppoint = RjzdList[0];
                            --RgcCount;
                        }
                        else
                        {
                            //  System.Windows.Forms.MessageBox.Show("筛选出的界址点与界址线错误！");
                            return false;
                        }
                    }

                    Rtmpline = getRelationLine(Rtmppoint.ShapeCopy, RjzxList);
                    if (Rtmpline == null)
                    {

                        //   System.Windows.Forms.MessageBox.Show("筛选出的界址点与界址线错误！");
                        return false;
                    }

                    RjzdSorted.Add(Rtmppoint);
                    RjzdList.Remove(Rtmppoint);
                    RjzxSorted.Add(Rtmpline);
                    RjzxList.Remove(Rtmpline);

                    Rtmppoint = getRelationPoint(Rtmpline.ShapeCopy, RjzdList);

                }

                //如果界址点数量大于九需要加方格数量错误！应该启下一页
                if (LjzxSorted.Count > 9)
                    createRows(sheetSource, 27, LjzxSorted.Count - 9);
                ///.......................................................
                int Li = 10;
                int LtmpInt = -1;
                bool Lfirst = true;
                foreach (IFeature Lf in LjzdSorted)
                {
                    if (Lf != null)
                    {
                        tmprow = sheetSource.GetRow(Li);
                        tmpcell = tmprow.GetCell(0);
                        tmpcell.SetCellValue(Lf.get_Value(jzdhIndex).ToString());

                        tmpstr = Lf.get_Value(jzdlxIndex).ToString();
                        if (jblxHT[tmpstr] == null)
                            LtmpInt = (int)jblxHT["null"];
                        else
                            LtmpInt = (int)jblxHT[tmpstr];

                        tmpcell = tmprow.GetCell(LtmpInt);
                        tmpcell.SetCellValue("√");
                    }
                    if (Lfirst)
                    {
                        Li += 1;
                        Lfirst = false;
                    }
                    else
                        Li += 2;
                }

                Li = 10;
                foreach (IFeature Lf in LjzxSorted)
                {
                    if (Lf != null)
                    {
                        tmprow = sheetSource.GetRow(Li);

                        tmpstr = Lf.get_Value(jzxlbIndex).ToString();
                        if (jzxlxHT[tmpstr] == null)
                            LtmpInt = (int)jzxlxHT["null"];
                        else
                            LtmpInt = (int)jzxlxHT[tmpstr];

                        tmpcell = tmprow.GetCell(LtmpInt);
                        tmpcell.SetCellValue("√");

                        tmpstr = Lf.get_Value(jzxwzIndex).ToString();
                        if (jzxwzHT[tmpstr] == null)
                            LtmpInt = (int)jzxwzHT["null"];
                        else
                            LtmpInt = (int)jzxwzHT[tmpstr];

                        tmpcell = tmprow.GetCell(LtmpInt);
                        tmpcell.SetCellValue("√");

                        tmpcell = tmprow.GetCell(18);
                        tmpcell.SetCellValue(Lf.get_Value(jzxsmIndex).ToString());

                        tmpcell = tmprow.GetCell(19);
                        tmpstr = Lf.get_Value(pldwqlrIndex).ToString();
                        string[] qlrArr = tmpstr.Split(',');
                        //左边0是承包方，那么权利人指界人都是1
                        if (qlrArr[0].CompareTo(cbfmc) == 0)
                        {
                            if (qlrArr[1].CompareTo("") == 0)
                            {
                                tmpcell.SetCellValue("/");
                                tmpcell = tmprow.GetCell(20);
                                tmpcell.SetCellValue(InfoHashTable["fbffzr"] + "");
                            }
                            else
                            {
                                tmpcell.SetCellValue(qlrArr[1]);
                                tmpcell = tmprow.GetCell(20);
                                tmpcell.SetCellValue(qlrArr[1]);
                            }
                        }
                        else//左边0不是承包方，那么权利人指界人都是0
                        {
                            if (qlrArr[0].CompareTo("") == 0)
                            {
                                tmpcell.SetCellValue("/");
                                tmpcell = tmprow.GetCell(20);
                                tmpcell.SetCellValue(InfoHashTable["fbffzr"] + "");
                            }
                            else
                            {
                                tmpcell.SetCellValue(qlrArr[0]);
                                tmpcell = tmprow.GetCell(20);
                                tmpcell.SetCellValue(qlrArr[0]);
                            }
                        }
                    }
                    Li += 2;
                }
                Li = Li > 25 ? Li + 1 : 26;
                //sheetSource.AddMergedRegion(new CellRangeAddress(10, Li, CommonHelper.Col('V'), CommonHelper.Col('V')));

                /////RRRRRRRRRRRRRRRRRRRRRRR
                if (RjzxSorted.Count > 9)
                    createRows(sheetSource, 27, RjzxSorted.Count - 9);
                int Ri = 10;
                int RtmpInt = -1;
                bool Rfirst = true;
                foreach (IFeature Rf in RjzdSorted)
                {
                    if (Rf != null)
                    {
                        tmprow = sheetSource.GetRow(Ri);
                        tmpcell = tmprow.GetCell(23);
                        tmpcell.SetCellValue(Rf.get_Value(jzdhIndex).ToString());

                        tmpstr = Rf.get_Value(jzdlxIndex).ToString();
                        if (jblxHT[tmpstr] == null)
                            RtmpInt = (int)jblxHT["null"];
                        else
                            RtmpInt = (int)jblxHT[tmpstr];

                        tmpcell = tmprow.GetCell(RtmpInt + 23);
                        tmpcell.SetCellValue("√");
                    }
                    if (Rfirst)
                    {
                        Ri += 1;
                        Rfirst = false;
                    }
                    else
                        Ri += 2;
                }

                Ri = 10;
                foreach (IFeature Rf in RjzxSorted)
                {
                    if (Rf != null)
                    {
                        tmprow = sheetSource.GetRow(Ri);

                        tmpstr = Rf.get_Value(jzxlbIndex).ToString();
                        if (jzxlxHT[tmpstr] == null)
                            RtmpInt = (int)jzxlxHT["null"];
                        else
                            RtmpInt = (int)jzxlxHT[tmpstr];

                        tmpcell = tmprow.GetCell(RtmpInt + 23);
                        tmpcell.SetCellValue("√");

                        tmpstr = Rf.get_Value(jzxwzIndex).ToString();
                        if (jzxwzHT[tmpstr] == null)
                            RtmpInt = (int)jzxwzHT["null"];
                        else
                            RtmpInt = (int)jzxwzHT[tmpstr];

                        tmpcell = tmprow.GetCell(RtmpInt + 23);
                        tmpcell.SetCellValue("√");

                        tmpcell = tmprow.GetCell(18 + 23);
                        tmpcell.SetCellValue(Rf.get_Value(jzxsmIndex).ToString());

                        tmpcell = tmprow.GetCell(19 + 23);
                        tmpstr = Rf.get_Value(pldwqlrIndex).ToString();
                        string[] qlrArr = tmpstr.Split(',');
                        //左边0是承包方，那么权利人指界人都是1
                        if (qlrArr[0].CompareTo(cbfmc) == 0)
                        {
                            if (qlrArr[1].CompareTo("") == 0)
                            {
                                tmpcell.SetCellValue("/");
                                tmpcell = tmprow.GetCell(20 + 23);
                                tmpcell.SetCellValue(InfoHashTable["fbffzr"] + "");
                            }
                            else
                            {
                                tmpcell.SetCellValue(qlrArr[1]);
                                tmpcell = tmprow.GetCell(20 + 23);
                                tmpcell.SetCellValue(qlrArr[1]);
                            }
                        }
                        else//左边0不是承包方，那么权利人指界人都是0
                        {
                            if (qlrArr[0].CompareTo("") == 0)
                            {
                                tmpcell.SetCellValue("/");
                                tmpcell = tmprow.GetCell(20 + 23);
                                tmpcell.SetCellValue(InfoHashTable["fbffzr"] + "");
                            }
                            else
                            {
                                tmpcell.SetCellValue(qlrArr[0]);
                                tmpcell = tmprow.GetCell(20 + 23);
                                tmpcell.SetCellValue(qlrArr[0]);
                            }
                        }
                    }
                    Ri += 2;
                }
                Ri = Ri > 25 ? Ri + 1 : 26;
                //sheetSource.AddMergedRegion(new CellRangeAddress(10, Ri, CommonHelper.Col('V'), CommonHelper.Col('V')));
                //本地块指阶人
                tmprow = sheetSource.GetRow(10);
                tmpcell = tmprow.GetCell(21);
                cbfmc = LzdtF.get_Value(LcbfmcIndex).ToString();
                tmpcell.SetCellValue(LzdtF.get_Value(LcbfmcIndex).ToString());

                tmpcell = tmprow.GetCell(21 + 23);
                cbfmc = RzdtF.get_Value(RcbfmcIndex).ToString();
                tmpcell.SetCellValue(RzdtF.get_Value(RcbfmcIndex).ToString());


                //lllllllllllllllllllllllll调查员
                tmprow = sheetSource.GetRow(31);
                tmpcell = tmprow.GetCell(4);
                tmpcell.SetCellValue(InfoHashTable["dcy"].ToString());
                //审核意见
                tmprow = sheetSource.GetRow(32);
                tmpcell = tmprow.GetCell(1);
                tmpcell.SetCellValue("合格");
                //日期两个
                tmprow = sheetSource.GetRow(31);
                tmpcell = tmprow.GetCell(18);
                DateTime cbfdcrq = Convert.ToDateTime(InfoHashTable["dcrq"].ToString());
                TimeSpan timeSpan = new TimeSpan(8, 0, 0, 0);
                DateTime gsshrq = cbfdcrq.Add(timeSpan);
                tmpcell.SetCellValue(cbfdcrq.ToLongDateString());

                tmprow = sheetSource.GetRow(34);
                tmpcell = tmprow.GetCell(18);
                tmpcell.SetCellValue(gsshrq.ToLongDateString());

                //rrrrrrrrrrrrr调查员
                tmprow = sheetSource.GetRow(31);
                tmpcell = tmprow.GetCell(4 + 23);
                tmpcell.SetCellValue(InfoHashTable["dcy"].ToString());
                //审核意见
                tmprow = sheetSource.GetRow(32);
                tmpcell = tmprow.GetCell(1 + 23);
                tmpcell.SetCellValue("合格");
                //日期两个

                tmprow = sheetSource.GetRow(31);
                tmpcell = tmprow.GetCell(18 + 23);
                tmpcell.SetCellValue(cbfdcrq.ToLongDateString());

                tmprow = sheetSource.GetRow(34);
                tmpcell = tmprow.GetCell(18 + 23);
                tmpcell.SetCellValue(gsshrq.ToLongDateString());
                //保存
                System.IO.FileStream fs = new System.IO.FileStream(outpath, FileMode.Open, FileAccess.ReadWrite);
                workbookSource.Write(fs);
                fs.Close();
                return true;
            }
            return false;
        }
        private bool CreateOneCTable(string databaseUrl, int fid, string outpath, string zdtname, string jzxname, string jzdname, Hashtable InfoHashTable)
        {
            try
            {
                Hashtable tempHashTable = InfoHashTable;
                IAeFactory pAeFactory = new PersonalGeoDatabase(databaseUrl);
                IFeatureWorkspace workspace = pAeFactory.OpenFeatrueWorkspace();
                IFeatureClass zdt = workspace.OpenFeatureClass(zdtname);
                IFeatureClass jzx = workspace.OpenFeatureClass(jzxname);
                IFeatureClass jzd = workspace.OpenFeatureClass(jzdname);

                var ret = CreateCTable(workspace, outpath, zdt.GetFeature(fid), jzx, jzd, tempHashTable);
                Marshal.FinalReleaseComObject(zdt);
                Marshal.FinalReleaseComObject(workspace);
                GC.WaitForPendingFinalizers();
                GC.Collect();
                return ret;
            }
            catch
            {
            }
            return false;
        }
        private bool CreateCTable(IFeatureWorkspace workspace, string outpath, IFeature zdtF, IFeatureClass jzx, IFeatureClass jzd, Hashtable InfoHashTable)
        {

            Hashtable tdytHashTable = new Hashtable();
            tdytHashTable.Add("1", "种植业");
            tdytHashTable.Add("2", "林业");
            tdytHashTable.Add("3", "畜牧业");
            tdytHashTable.Add("4", "渔业");
            tdytHashTable.Add("5", "其他");

            Hashtable tdlylxHashTable = new Hashtable();
            tdlylxHashTable.Add("011", "水田");
            tdlylxHashTable.Add("012", "水浇地");
            tdlylxHashTable.Add("013", "旱地");

            Hashtable dldjHashTable = new Hashtable();
            dldjHashTable.Add("01", "一等地");
            dldjHashTable.Add("02", "二等地");
            dldjHashTable.Add("03", "三等地");
            dldjHashTable.Add("04", "四等地");
            dldjHashTable.Add("05", "五等地");
            dldjHashTable.Add("06", "六等地");
            dldjHashTable.Add("07", "七等地");

            Hashtable jblxHT = new Hashtable();
            jblxHT.Add("1", 1);
            jblxHT.Add("2", 2);
            jblxHT.Add("3", 3);
            jblxHT.Add("4", 4);
            jblxHT.Add("null", 5);

            Hashtable jzxlxHT = new Hashtable();
            jzxlxHT.Add("01", 6);
            jzxlxHT.Add("02", 7);
            jzxlxHT.Add("03", 8);
            jzxlxHT.Add("04", 9);
            jzxlxHT.Add("05", 10);
            jzxlxHT.Add("06", 11);
            jzxlxHT.Add("07", 12);
            jzxlxHT.Add("08", 13);
            jzxlxHT.Add("09", 14);

            Hashtable jzxwzHT = new Hashtable();
            jzxwzHT.Add("1", 15);
            jzxwzHT.Add("2", 16);
            jzxwzHT.Add("3", 17);



            string templateUrl = AppDomain.CurrentDomain.BaseDirectory + @"template\承包地块调查表.xls";


            int fbfbmIndex = zdtF.Fields.FindField("FBFBM");
            int cbfmcIndex = zdtF.Fields.FindField("CBFMC");
            int dkbhIndex = zdtF.Fields.FindField("DKBM");
            int dkmcIndex = zdtF.Fields.FindField("DKMC");
            int htmjIndex = zdtF.Fields.FindField("HTMJ");
            int dzIndex = zdtF.Fields.FindField("DKDZ");
            int xzIndex = zdtF.Fields.FindField("DKXZ");
            int nzIndex = zdtF.Fields.FindField("DKNZ");
            int bzIndex = zdtF.Fields.FindField("DKBZ");
            int tdytIndex = zdtF.Fields.FindField("TDYT");
            int tdlylxIndex = zdtF.Fields.FindField("TDLYLX");
            int dldjIndex = zdtF.Fields.FindField("DLDJ");
            int sfjbntIndex = zdtF.Fields.FindField("SFJBNT");
            int jzdlxIndex = jzd.Fields.FindField("JBLX");
            int jzdhIndex = jzd.Fields.FindField("JZDH");
            int jzxlbIndex = jzx.Fields.FindField("JZXLB");
            int jzxwzIndex = jzx.Fields.FindField("JZXWZ");
            int jzxsmIndex = jzx.Fields.FindField("JZXSM");
            int pldwqlrIndex = jzx.Fields.FindField("PLDWQLR");
            int pldwzjrIndex = jzx.Fields.FindField("PLDWZJR");

            // if (File.Exists(outpath)) File.Delete(outpath);
            File.Copy(templateUrl, outpath, true);
            using (System.IO.FileStream fileStream = new System.IO.FileStream(outpath, FileMode.Open, FileAccess.ReadWrite))
            {
                HSSFWorkbook workbookSource = new HSSFWorkbook(fileStream);
                HSSFSheet sheetSource = (HSSFSheet)workbookSource.GetSheetAt(0);
                //设定合并单元格的样式
                HSSFCellStyle style = (HSSFCellStyle)workbookSource.CreateCellStyle();
                style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.CENTER;
                style.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.CENTER;
                style.BorderBottom = NPOI.SS.UserModel.BorderStyle.THIN;
                style.BorderRight = NPOI.SS.UserModel.BorderStyle.THIN;
                style.BorderLeft = NPOI.SS.UserModel.BorderStyle.THIN;
                style.BorderTop = NPOI.SS.UserModel.BorderStyle.THIN;
                style.WrapText = true;

                string cbfmc = null;

                var tmprow = sheetSource.GetRow(2);
                var tmpcell = tmprow.GetCell(4);
                tmpcell.SetCellValue(zdtF.get_Value(fbfbmIndex).ToString());

                tmpcell = tmprow.GetCell(12);
                cbfmc = zdtF.get_Value(cbfmcIndex).ToString();
                tmpcell.SetCellValue(zdtF.get_Value(cbfmcIndex).ToString());

                tmprow = sheetSource.GetRow(3);
                tmpcell = tmprow.GetCell(4);
                tmpcell.SetCellValue(zdtF.get_Value(dkbhIndex).ToString().Substring(14));

                tmpcell = tmprow.GetCell(12);
                tmpcell.SetCellValue(zdtF.get_Value(dkmcIndex).ToString());

                tmpcell = tmprow.GetCell(20);
                tmpcell.SetCellValue(((double)zdtF.get_Value(htmjIndex)).ToString("f"));

                tmprow = sheetSource.GetRow(4);
                tmpcell = tmprow.GetCell(1);
                tmpcell.SetCellValue(zdtF.get_Value(dzIndex).ToString());

                tmpcell = tmprow.GetCell(9);
                tmpcell.SetCellValue(zdtF.get_Value(nzIndex).ToString());

                tmpcell = tmprow.GetCell(16);
                tmpcell.SetCellValue(zdtF.get_Value(xzIndex).ToString());

                tmpcell = tmprow.GetCell(20);
                tmpcell.SetCellValue(zdtF.get_Value(bzIndex).ToString());

                tmprow = sheetSource.GetRow(5);
                tmpcell = tmprow.GetCell(1);
                string tmpstr = tdytHashTable[zdtF.get_Value(tdytIndex)].ToString();
                IRichTextString rich = tmpcell.RichStringCellValue;
                IFont font = workbookSource.GetFontAt(rich.GetFontAtIndex(rich.String.Length - 1));
                tmpstr = rich.String.Replace("□" + tmpstr, "■" + tmpstr);
                rich = new HSSFRichTextString(tmpstr);
                rich.ApplyFont(tmpstr.IndexOf('_'), tmpstr.LastIndexOf('_'), font);
                tmpcell.SetCellValue(rich);

                tmpcell = tmprow.GetCell(15);
                tmpstr = tdlylxHashTable[zdtF.get_Value(tdlylxIndex)].ToString();
                rich = tmpcell.RichStringCellValue;
                font = workbookSource.GetFontAt(rich.GetFontAtIndex(rich.String.Length - 1));
                tmpstr = rich.String.Replace("□" + tmpstr, "■" + tmpstr);
                rich = new HSSFRichTextString(tmpstr);
                rich.ApplyFont(tmpstr.IndexOf('_'), tmpstr.LastIndexOf('_'), font);
                tmpcell.SetCellValue(rich);

                tmpcell = tmprow.GetCell(10);
                tmpstr = dldjHashTable[zdtF.get_Value(dldjIndex)].ToString();
                tmpcell.SetCellValue(tmpstr);

                tmpcell = tmprow.GetCell(21);
                tmpstr = zdtF.get_Value(sfjbntIndex).ToString();
                if ("2".CompareTo(tmpstr) != 0)
                {
                    tmpcell.SetCellValue("■是");
                }
                else
                {
                    tmprow = sheetSource.GetRow(6);
                    tmpcell = tmprow.GetCell(21);
                    tmpcell.SetCellValue("■否");
                }

                ISpatialFilter sf = new SpatialFilterClass();
                sf.Geometry = zdtF.ShapeCopy;
                sf.GeometryField = "SHAPE";
                sf.SpatialRel = esriSpatialRelEnum.esriSpatialRelRelation;
                sf.SpatialRelDescription = "F*TT*TF*T";
                var cursor = jzd.Search(sf, false);
                IFeature tmpFeature = null;
                List<IFeature> jzdList = new List<IFeature>();
                while ((tmpFeature = cursor.NextFeature()) != null)
                {
                    jzdList.Add(tmpFeature);
                }
                IGeometryCollection gc = zdtF.ShapeCopy as IGeometryCollection;
                if (gc.GeometryCount > 1)
                {
                    int a = gc.GeometryCount;
                }
                sf = new SpatialFilterClass();
                sf.GeometryField = "SHAPE";
                sf.Geometry = zdtF.ShapeCopy;
                sf.SpatialRel = esriSpatialRelEnum.esriSpatialRelRelation;
                sf.SpatialRelDescription = "FFTTT*FF*";

                cursor = jzx.Search(sf, false);
                tmpFeature = null;
                List<IFeature> jzxList = new List<IFeature>();

                while ((tmpFeature = cursor.NextFeature()) != null)
                {
                    jzxList.Add(tmpFeature);
                }

                List<IFeature> jzdSorted = new List<IFeature>();
                List<IFeature> jzxSorted = new List<IFeature>();

                if (jzdList.Count != jzxList.Count)
                {
                    //System.Windows.Forms.MessageBox.Show("筛选出的界址点与界址线数量不同！");
                    return false;
                }

                int j = jzdList.Count;
                IFeature tmppoint = jzdList[0];
                IFeature tmpline = null;
                int gcCount = (zdtF.ShapeCopy as IGeometryCollection).GeometryCount - 1;
                while (jzxList.Count > 0)
                {
                    if (tmppoint == null)
                    {
                        if (gcCount > 0)
                        {
                            jzxSorted.Add(null);
                            jzdSorted.Add(null);
                            tmppoint = jzdList[0];
                            --gcCount;
                        }
                        else
                        {
                            //System.Windows.Forms.MessageBox.Show("筛选出的界址点与界址线错误！");
                            return false;
                        }
                    }

                    tmpline = getRelationLine(tmppoint.ShapeCopy, jzxList);
                    if (tmpline == null)
                    {

                        //  System.Windows.Forms.MessageBox.Show("筛选出的界址点与界址线错误！");
                        return false;
                    }

                    jzdSorted.Add(tmppoint);
                    jzdList.Remove(tmppoint);
                    jzxSorted.Add(tmpline);
                    jzxList.Remove(tmpline);

                    tmppoint = getRelationPoint(tmpline.ShapeCopy, jzdList);

                }

                if (jzxSorted.Count > 9)
                    createRows(sheetSource, 27, jzxSorted.Count - 9);
                int i = 10;
                int tmpInt = -1;
                bool first = true;
                foreach (IFeature f in jzdSorted)
                {
                    if (f != null)
                    {
                        tmprow = sheetSource.GetRow(i);
                        tmpcell = tmprow.GetCell(0);
                        tmpcell.SetCellValue(f.get_Value(jzdhIndex).ToString());

                        tmpstr = f.get_Value(jzdlxIndex).ToString();
                        if (jblxHT[tmpstr] == null)
                            tmpInt = (int)jblxHT["null"];
                        else
                            tmpInt = (int)jblxHT[tmpstr];

                        tmpcell = tmprow.GetCell(tmpInt);
                        tmpcell.SetCellValue("√");
                    }
                    if (first)
                    {
                        i += 1;
                        first = false;
                    }
                    else
                        i += 2;
                }

                i = 10;
                //、、、、、、、、、、、、、、、、、、、、、、、、、、、、、、、、、、、、、、、、、
                foreach (IFeature f in jzxSorted)
                {
                    if (f != null)
                    {
                        //拿到当前界址线编辑行行号为i
                        tmprow = sheetSource.GetRow(i);
                        //拿出HT中对应的界址线类别-->tmpInt
                        tmpstr = f.get_Value(jzxlbIndex).ToString();
                        if (jzxlxHT[tmpstr] == null)
                            tmpInt = (int)jzxlxHT["null"];
                        else
                            tmpInt = (int)jzxlxHT[tmpstr];

                        tmpcell = tmprow.GetCell(tmpInt);
                        tmpcell.SetCellValue("√");
                        //拿出HT中对应的界址线位置-->tmpInt为列数
                        tmpstr = f.get_Value(jzxwzIndex).ToString();
                        if (jzxwzHT[tmpstr] == null)
                            tmpInt = (int)jzxwzHT["null"];
                        else
                            tmpInt = (int)jzxwzHT[tmpstr];

                        tmpcell = tmprow.GetCell(tmpInt);
                        tmpcell.SetCellValue("√");

                        tmpcell = tmprow.GetCell(18);
                        tmpcell.SetCellValue(f.get_Value(jzxsmIndex).ToString());

                        //19权利人，20指界人
                        tmpcell = tmprow.GetCell(19);
                        tmpstr = f.get_Value(pldwqlrIndex).ToString();
                        string[] qlrArr = tmpstr.Split(',');
                        //左边0是承包方，那么权利人指界人都是1
                        if (qlrArr[0].CompareTo(cbfmc) == 0)
                        {
                            if (qlrArr[1].CompareTo("") == 0)
                            {
                                tmpcell.SetCellValue("/");
                                tmpcell = tmprow.GetCell(20);
                                tmpcell.SetCellValue(InfoHashTable["fbffzr"] + "");
                            }
                            else
                            {
                                tmpcell.SetCellValue(qlrArr[1]);
                                tmpcell = tmprow.GetCell(20);
                                tmpcell.SetCellValue(qlrArr[1]);
                            }
                        }
                        else//左边0不是承包方，那么权利人指界人都是0
                        {
                            if (qlrArr[0].CompareTo("") == 0)
                            {
                                tmpcell.SetCellValue("/");
                                tmpcell = tmprow.GetCell(20);
                                tmpcell.SetCellValue(InfoHashTable["fbffzr"] + "");
                            }
                            else
                            {
                                tmpcell.SetCellValue(qlrArr[0]);
                                tmpcell = tmprow.GetCell(20);
                                tmpcell.SetCellValue(qlrArr[0]);
                            }
                        }
                    }
                    i += 2;
                }
                i = i > 25 ? i + 1 : 26;
                //本地块指阶人
                tmprow = sheetSource.GetRow(10);
                tmpcell = tmprow.GetCell(21);
                cbfmc = zdtF.get_Value(cbfmcIndex).ToString();
                tmpcell.SetCellValue(zdtF.get_Value(cbfmcIndex).ToString());
                //调查员
                tmprow = sheetSource.GetRow(31);
                tmpcell = tmprow.GetCell(4);
                tmpcell.SetCellValue(InfoHashTable["dcy"].ToString());
                //审核意见
                tmprow = sheetSource.GetRow(32);
                tmpcell = tmprow.GetCell(1);
                tmpcell.SetCellValue("合格");
                //日期两个
                DateTime cbfdcrq = Convert.ToDateTime(InfoHashTable["dcrq"].ToString());
                TimeSpan timeSpan = new TimeSpan(8, 0, 0, 0);
                DateTime gsshrq = cbfdcrq.Add(timeSpan);
                tmprow = sheetSource.GetRow(31);
                tmpcell = tmprow.GetCell(18);
                tmpcell.SetCellValue(cbfdcrq.ToLongDateString());

                tmprow = sheetSource.GetRow(34);
                tmpcell = tmprow.GetCell(18);
                tmpcell.SetCellValue(gsshrq.ToLongDateString());

                //sheetSource.AddMergedRegion(new CellRangeAddress(10, i, CommonHelper.Col('V'), CommonHelper.Col('V')));
                System.IO.FileStream fs = new System.IO.FileStream(outpath, FileMode.Open, FileAccess.ReadWrite);
                workbookSource.Write(fs);
                fs.Close();
                return true;
            }
            return false;
        }
        public void createRows(HSSFSheet sourceSheet, int rownum, int count)
        {
            sourceSheet.ShiftRows(27, 34, count * 2, true, false, true);
            for (int i = 0; i < count; i++)
            {
                createTwoRow(sourceSheet, rownum + 2 * i);
            }
        }
        private void createTwoRow(HSSFSheet sourceSheet, int rownum)
        {
            createOneRow(sourceSheet, rownum);
            createOneRow(sourceSheet, rownum + 1);
            for (int i = 0; i < 6; i++)
            {
                sourceSheet.AddMergedRegion(new CellRangeAddress(rownum, rownum + 1, i, i));
                sourceSheet.SetEnclosedBorderOfRegion(new CellRangeAddress(rownum, rownum + 1, i, i), NPOI.SS.UserModel.BorderStyle.THIN, NPOI.HSSF.Util.HSSFColor.BLACK.index);
            }
            for (int i = 6; i < 22; i++)
            {
                sourceSheet.AddMergedRegion(new CellRangeAddress(rownum - 1, rownum, i, i));
                sourceSheet.SetEnclosedBorderOfRegion(new CellRangeAddress(rownum - 1, rownum, i, i), NPOI.SS.UserModel.BorderStyle.THIN, NPOI.HSSF.Util.HSSFColor.BLACK.index);
                sourceSheet.SetEnclosedBorderOfRegion(new CellRangeAddress(rownum + 1, rownum + 1, i, i), NPOI.SS.UserModel.BorderStyle.THIN, NPOI.HSSF.Util.HSSFColor.BLACK.index);
            }


        }
        private void createOneRow(HSSFSheet sourceSheet, int rownum)
        {
            var srcRow = sourceSheet.GetRow(rownum - 1);
            var newRow = sourceSheet.CreateRow(rownum);
            newRow.Height = srcRow.Height;

            for (int m = 0; m < 22; m++)
            {
                ICell cell = newRow.CreateCell(m);
                // cell.SetCellValue("0");
                ICellStyle cellStyle = srcRow.Cells[m].CellStyle;

                cell.CellStyle = cellStyle;
                cell.SetCellType(newRow.Cells[m].CellType);
            }
        }
    }
}
