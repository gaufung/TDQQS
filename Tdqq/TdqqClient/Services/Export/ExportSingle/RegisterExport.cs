﻿using System;
using System.Collections.Generic;
using Aspose.Words;
using TdqqClient.Services.Common;

namespace TdqqClient.Services.Export.ExportSingle
{
    class RegisterExport:ExportBase
    {
        public RegisterExport(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        { }

        public override void Export(string cbfmc, string cbfbm, string folderPath, string edgeFeature = "")
        {
            var saveFilePath = folderPath + @"\" + cbfbm.Substring(14) + "_" + cbfmc + "_08承包经营权证登记簿.pdf";
            FillDocFile(saveFilePath, cbfbm);
        }

        private void FillDocFile(string savedFilePath, string cbfbm)
        {
            var fbf = Fbf();
            var cbf = Cbf(cbfbm);
            var cbfJtcy = Cbfjtcys(cbfbm);
            var fields = Fields(cbfbm);
            //如果获得数据不符合规范
            if (fbf == null || cbf == null || cbfJtcy == null || fields == null) return;
            //获取登记簿模板的位置
            var fileDocSourth = AppDomain.CurrentDomain.BaseDirectory + @"\template\登记薄" + fields.Count +".doc";
            Document doc = new Document(fileDocSourth);
            DocumentBuilder documentBuilder = new DocumentBuilder(doc);
            //书签处插入
            doc.Range.Bookmarks["经营权证号"].Text = cbfbm + "J";
            doc.Range.Bookmarks["fbfmc"].Text = fbf.Fbfmc;
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
            documentBuilder.MoveToSection(1);
            double fontSize = 0.0;
            for (int i = 0; i < columnCount; i++)
            {
                documentBuilder.MoveToCell(0, 9, i, 0); //
                double width = documentBuilder.CellFormat.Width; //获取单元格宽度
                fontSize = documentBuilder.Font.Size;
                widthList.Add(width);
            }
            //插入承包方表格
            documentBuilder.MoveToBookmark("table");
            for (int i = 0; i < cbfJtcy.Count; i++)
            {
                for (int j = 0; j < columnCount; j++)
                {
                    documentBuilder.InsertCell(); // 添加一个单元格                    
                    documentBuilder.CellFormat.Borders.LineStyle = Aspose.Words.LineStyle.Single;
                    documentBuilder.CellFormat.Borders.Color = System.Drawing.Color.Black;
                    documentBuilder.CellFormat.Width = widthList[j];
                    documentBuilder.Font.Size = fontSize;
                    documentBuilder.Font.Name = "宋体";
                    documentBuilder.CellFormat.VerticalMerge = Aspose.Words.Tables.CellMerge.None;
                    documentBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                    /*
                     * 承包方家庭成员信息：CBFBM,CBFMC,CYXB,CYXM,CYZJHM,CYZJLX,CYBZ,CYBZ,YHZGX,CYSZC,YZBM,LXDH,SFGYR
                     */
                    if (j == 0)
                    {
                        documentBuilder.Write(cbfJtcy[i].Cyxm);
                    }
                    if (j == 1)
                    {
                        documentBuilder.Write(Transcode.CodeToRelationship(cbfJtcy[i].Yhzgx));
                    }
                    if (j == 2)
                    {
                        documentBuilder.Write(cbfJtcy[i].Cyzjhm);
                    }
                    if (j == 3)
                    {
                        documentBuilder.Write(cbfJtcy[i].Cybz);
                    }
                }
                documentBuilder.EndRow();
            }
            double htSum = 0.0;
            double scSum = 0.0;
            //插入地块信息
            /*
             *  CBFMC,DKMC,YHTMJ,DKBM,DKDZ,DKNZ,DKXZ,DKBZ,DKBZXX,ZJRXM,DKLB,TDLYLX,DLDJ,TDYT,SFJBNT,CBJYQQDFS,HTMJ,SCMJ
             */
            for (int i = 0; i < fields.Count; i++)
            {
                doc.Range.Bookmarks["mc" + (i + 1)].Text = fields[i].Dkmc;
                doc.Range.Bookmarks["bm" + (i + 1)].Text = fields[i].Dkbm;
                scSum += fields[i].Scmj;
                htSum += fields[i].Htmj;
                doc.Range.Bookmarks["ht" + (i + 1)].Text = fields[i].Htmj.ToString("f");
                doc.Range.Bookmarks["sc" + (i + 1)].Text = fields[i].Scmj.ToString("f");
                doc.Range.Bookmarks["sf" + (i + 1)].Text =
                    Transcode.CodeToSfjbnt(fields[i].Sfjbnt);
                doc.Range.Bookmarks["d" + (i + 1)].Text = EditSz(fields[i].Dkdz);
                doc.Range.Bookmarks["n" + (i + 1)].Text = EditSz(fields[i].Dknz);
                doc.Range.Bookmarks["x" + (i + 1)].Text = EditSz(fields[i].Dkxz);
                doc.Range.Bookmarks["b" + (i + 1)].Text = EditSz(fields[i].Dkbz);
            }
            doc.Range.Bookmarks["htsum"].Text = htSum.ToString("f");
            doc.Range.Bookmarks["scsum"].Text = scSum.ToString("f");
            doc.Range.Bookmarks["dkzs"].Text = fields.Count.ToString();
            doc.Save(savedFilePath, SaveFormat.Pdf);

        }
    }
}
