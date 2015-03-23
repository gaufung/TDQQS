using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using TdqqClient.Services.Common;

namespace TdqqClient.Services.Export.ExportSingle
{
    class JyqzExport:ExportBase
    {

        public JyqzExport(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        { }

        /// <summary>
        /// 导出经营权证
        /// </summary>
        /// <param name="cbfmc">承包方名称</param>
        /// <param name="cbfbm">承包方编码</param>
        /// <param name="folderPath">文件夹地址</param>
        /// <param name="edgeFeature">边界要素类</param>
        public override void Export(string cbfmc, string cbfbm, string folderPath, string edgeFeature = "")
        {
            var dir = new DirectoryInfo(folderPath);
            //创造子文件夹
            dir.CreateSubdirectory(cbfbm.Substring(14) + cbfmc);
            var subFolder = Path.Combine(folderPath, cbfbm.Substring(14, 4) + cbfmc);
            var saveDocPath = subFolder + @"\" + cbfbm.Substring(14) + "_" + cbfmc + "_经营权证.doc";
            var saveAppendixPath = subFolder + @"\" + cbfbm.Substring(14) + "_" + cbfmc + "_承包方信息附录.pdf";
            ExportCertifiCation(saveDocPath, cbfbm,cbfmc);
            ExportAppendix(saveAppendixPath, cbfbm);
        }
        /// <summary>
        /// 导出证书信息
        /// </summary>
        /// <param name="toSaveFilePath">保存文件的地址</param>
        /// <param name="cbfbm">承包方编码</param>
        private void ExportCertifiCation(string toSaveFilePath, string cbfbm,string cbfmc)
        {
  
            var cbf = Cbf(cbfbm);
            if (cbf == null) return;
            var fields = Fields(cbfbm);          
            var cbfzjhm =cbf.Cbfzjhm;
            var yzbm = cbf.Yzbm;
            var lxdh =cbf.Lxdh;
            var dcsh = DcSh();
            var year = dcsh.Gsshrq.Year.ToString();
            var month = dcsh.Gsshrq.Month.ToString();
            var day = dcsh.Gsshrq.Day.ToString();
            var fbf =Fbf();
            var fbfmc = fbf.Fbfmc.Substring(fbf.Fbfmc.IndexOf("县") + 1);
            string xian, zhen, cun;
            xian = zhen = cun = string.Empty;
            SpiltFbfdzToXianXiangCun(fbf.Fbfdz, ref xian, ref zhen, ref cun);
            //
            var templatePath = AppDomain.CurrentDomain.BaseDirectory + @"\template\承包经营权.doc";

            Document exportWord = new Document(templatePath);
            //填充封面
            exportWord.Range.Bookmarks["证书编号1"].Text = cbfbm + "J";
            exportWord.Range.Bookmarks["年1"].Text = year;
            exportWord.Range.Bookmarks["月1"].Text = month;
            exportWord.Range.Bookmarks["日1"].Text = day;
            exportWord.Range.Bookmarks["年2"].Text = year;
            exportWord.Range.Bookmarks["月2"].Text = month;
            exportWord.Range.Bookmarks["日2"].Text = day;
            //填充共有信息
            exportWord.Range.Bookmarks["发包方"].Text = fbfmc;
            exportWord.Range.Bookmarks["承包方名称"].Text = cbfmc;
            exportWord.Range.Bookmarks["身份证号码"].Text = cbfzjhm;
            exportWord.Range.Bookmarks["县"].Text = xian;
            exportWord.Range.Bookmarks["镇"].Text = zhen;
            exportWord.Range.Bookmarks["村"].Text = cun;
            exportWord.Range.Bookmarks["邮编"].Text = yzbm;
            exportWord.Range.Bookmarks["联系电话"].Text = lxdh;
            exportWord.Range.Bookmarks["证书编号2"].Text = cbfbm + "J";
            //根据每个田块填充
           
            var query = from field in fields
                select field.Scmj;
            double sumScmj = query.Sum();
            int fieldCount = fields.Count;            
            exportWord.Range.Bookmarks["实测面积"].Text = sumScmj.ToString("f");
            exportWord.Range.Bookmarks["地块数"].Text = fieldCount.ToString();
            for (int i = 0; i < fieldCount; i++)
            {
                exportWord.Range.Bookmarks["地块名称" + i].Text = fields[i].Dkmc;
                exportWord.Range.Bookmarks["地块编码" + i].Text = fields[i].Dkbm;
                exportWord.Range.Bookmarks["实测面积" + i].Text = fields[i].Scmj.ToString("F");

                exportWord.Range.Bookmarks["基本农田" + i].Text = Transcode.CodeToSfjbnt(fields[i].Sfjbnt);
                exportWord.Range.Bookmarks["东" + i].Text = EditSz(fields[i].Dkdz) + "\r\n";
                exportWord.Range.Bookmarks["西" + i].Text = EditSz(fields[i].Dkxz) + "\r\n";
                exportWord.Range.Bookmarks["南" + i].Text = EditSz(fields[i].Dknz) + "\r\n";
                exportWord.Range.Bookmarks["北" + i].Text = EditSz(fields[i].Dkbz);
            }
            for (int i = fieldCount; i < 17; i++)
            {
                exportWord.Range.Bookmarks["基本农田" + i].Text = " ";
                exportWord.Range.Bookmarks["地块名称" + i].Text = " ";
                exportWord.Range.Bookmarks["地块编码" + i].Text = " ";
                exportWord.Range.Bookmarks["实测面积" + i].Text = " ";
                exportWord.Range.Bookmarks["东" + i].Text = " ";
                exportWord.Range.Bookmarks["西" + i].Text = " ";
                exportWord.Range.Bookmarks["南" + i].Text = " ";
                exportWord.Range.Bookmarks["北" + i].Text = " ";
            }
            exportWord.Save(toSaveFilePath);
        }

        /// <summary>
        /// 导出附录信息
        /// </summary>
        /// <param name="toSaveFilePath">保存的文件地址</param>
        /// <param name="cbfbm">承包方编码</param>
        private void ExportAppendix(string toSaveFilePath, string cbfbm)
        {
            var fbf = Fbf();
            var cbf = Cbf(cbfbm);
            var cbfjtcys = Cbfjtcys(cbfbm);
            var fileDocSourth = AppDomain.CurrentDomain.BaseDirectory + @"\template\证书附录.doc";
            Document doc = new Document(fileDocSourth);
            DocumentBuilder documentBuilder = new DocumentBuilder(doc);
            //书签处插入
            doc.Range.Bookmarks["经营权证号"].Text = cbfbm + "J";
            doc.Range.Bookmarks["fbfmc"].Text =fbf.Fbfdz;
            doc.Range.Bookmarks["cbfmc"].Text = cbf.Cbfmc;
            doc.Range.Bookmarks["sfzhm"].Text = cbf.Cbfzjhm;
            doc.Range.Bookmarks["cbfdz"].Text = fbf.Fbfdz;
            doc.Range.Bookmarks["yzbm"].Text = cbf.Yzbm;
            doc.Range.Bookmarks["lxdh"].Text = cbf.Lxdh;
            doc.Range.Bookmarks["htbm"].Text = cbfbm + "J";
            List<double> widthList = new List<double>();
            //获取列宽
            const int columnCount = 4;
            //移动到该页
            documentBuilder.MoveToSection(0);
            double fontSize = 0.0;
            for (int i = 0; i < columnCount; i++)
            {
                documentBuilder.MoveToCell(0, 9, i, 0); //
                double width = documentBuilder.CellFormat.Width;//获取单元格宽度
                fontSize = documentBuilder.Font.Size;
                widthList.Add(width);
            }
            //插入承包方表格
            documentBuilder.MoveToBookmark("table");
            for (int i = 0; i < cbfjtcys.Count; i++)
            {              
                for (int j = 0; j < columnCount; j++)
                {
                    documentBuilder.InsertCell();            // 添加一个单元格                    
                    documentBuilder.CellFormat.Borders.LineStyle = Aspose.Words.LineStyle.Single;
                    documentBuilder.CellFormat.Borders.Color = System.Drawing.Color.Black;
                    documentBuilder.CellFormat.Width = widthList[j];
                    documentBuilder.Font.Size = fontSize;
                    documentBuilder.Font.Name = "宋体";
                    documentBuilder.CellFormat.VerticalMerge = Aspose.Words.Tables.CellMerge.None;
                    documentBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                    /*
                     *  CBFBM,CBFMC,CYXB,CYXM,CYZJHM,CYZJLX,CYBZ,CYBZ,YHZGX,CYSZC,YZBM,LXDH,SFGYR
                     *  0               1           2           3           4           5           6       7           8           9       10      11      12
                     */
                    if (j == 0)
                    {
                        documentBuilder.Write(cbfjtcys[i].Cyxm);
                    }
                    if (j == 1)
                    {
                        documentBuilder.Write(Transcode.CodeToRelationship(cbfjtcys[i].Yhzgx));
                    }
                    if (j == 2)
                    {
                        documentBuilder.Write(cbfjtcys[i].Cyzjhm);
                    }
                    if (j == 3)
                    {
                        documentBuilder.Write(cbfjtcys[i].Cybz);
                    }
                }
                documentBuilder.EndRow();
            }
            doc.Save(toSaveFilePath, SaveFormat.Pdf);
        }
        private static void SpiltFbfdzToXianXiangCun(string fbfdz, ref string xian, ref string zhen, ref string cun)
        {
            int xianIndex = fbfdz.IndexOf("县");
            int zhenIndex = fbfdz.IndexOf("镇") == -1 ? fbfdz.IndexOf("乡") : fbfdz.IndexOf("镇");
            int cunIndex = fbfdz.IndexOf("村");
            xian = fbfdz.Substring(0, xianIndex);
            zhen = fbfdz.Substring(xianIndex + 1, zhenIndex - xianIndex - 1);
            cun = fbfdz.Substring(zhenIndex + 1, cunIndex - zhenIndex - 1);
        }
    }
}
