using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using NPOI.SS.Formula.Functions;
using TdqqClient.Services.Common;

namespace TdqqClient.Models.Export.ExportSingle
{
    class RegisterExport:ExportBase
    {
        public RegisterExport(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }

        public override void Export(string cbfmc, string cbfbm, string folderPath, string edgeFeature = "")
        {
            var saveFilePath = folderPath + @"\" + cbfbm.Substring(14) + "_" + cbfmc + "_08承包经营权证登记簿.pdf";
            FillDocFile(saveFilePath,cbfbm);
        }

        private void FillDocFile(string savedFilePath, string cbfbm)
        {
            var rowFbf = SelectFbfInfo();
            var rowCbf = SelectCbfInfoByCbfbm(cbfbm);
            var dtCbfJtcy = SelectCbf_JtcyByCbfbm(cbfbm);
            var dtField = SelectFieldsByCbfbm(cbfbm);
            //如果获得数据不符合规范
            if (rowFbf == null || rowCbf == null || dtCbfJtcy == null || dtField == null) return;
            //获取登记簿模板的位置
            var fileDocSourth = AppDomain.CurrentDomain.BaseDirectory + @"\template\登记薄" + dtField.Rows.Count.ToString() +
                                ".doc";
            Document doc = new Document(fileDocSourth);
            DocumentBuilder documentBuilder = new DocumentBuilder(doc);
            //书签处插入
            /*
             * 发包方：select FBFBM,FBFMC,FBFFZRXM,FZRZJLX,FZRZJHM,LXDH,FBFDZ,YZBM,FBFDCY,FBFDCRQ,FBFDCJS
             * 承包方：CBFBM,CBFLX,CBFMC,CYXB,CBFZJLX,CBFZJHM,CBFDZ,YZBM,LXDH,CBFCYSL
             */
            doc.Range.Bookmarks["经营权证号"].Text = cbfbm + "J";
            doc.Range.Bookmarks["fbfmc"].Text = rowFbf[1].ToString();
            doc.Range.Bookmarks["cbfmc"].Text = rowCbf[2].ToString();
            doc.Range.Bookmarks["sfzhm"].Text = rowCbf[5].ToString();
            doc.Range.Bookmarks["cbfdz"].Text = rowFbf[6].ToString();
            doc.Range.Bookmarks["yzbm"].Text = rowCbf[7].ToString();
            doc.Range.Bookmarks["lxdh"].Text = rowCbf[8].ToString();
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
            for (int i = 0; i < dtCbfJtcy.Rows.Count; i++)
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
                        documentBuilder.Write(dtCbfJtcy.Rows[i][3].ToString());
                    }
                    if (j == 1)
                    {
                        documentBuilder.Write(Transcode.CodeToRelationship(dtCbfJtcy.Rows[i][8].ToString()));
                    }
                    if (j == 2)
                    {
                        documentBuilder.Write(dtCbfJtcy.Rows[i][4].ToString());
                    }
                    if (j == 3)
                    {
                        documentBuilder.Write(dtCbfJtcy.Rows[i][7].ToString());
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
            for (int i = 0; i < dtField.Rows.Count; i++)
            {
                doc.Range.Bookmarks["mc" + (i + 1).ToString()].Text = dtField.Rows[i][1].ToString();
                doc.Range.Bookmarks["bm" + (i + 1).ToString()].Text = dtField.Rows[i][3].ToString();
                double htSingle, scSingle;
                if (string.IsNullOrEmpty(dtField.Rows[i][16].ToString().Trim()))
                {
                    htSingle = 0.0;
                }
                else
                {
                    htSingle = Convert.ToDouble((Convert.ToDouble(dtField.Rows[i][16].ToString())).ToString("f"));
                }
                htSum += htSingle;
                if (string.IsNullOrEmpty(dtField.Rows[i][17].ToString()))
                {
                    scSingle = 0.0;
                }
                else
                {
                    scSingle = Convert.ToDouble((Convert.ToDouble(dtField.Rows[i][17].ToString())).ToString("f"));
                }
                scSum += scSingle;
                doc.Range.Bookmarks["ht" + (i + 1).ToString()].Text = htSingle.ToString("f");
                doc.Range.Bookmarks["sc" + (i + 1).ToString()].Text = scSingle.ToString("f");
                doc.Range.Bookmarks["sf" + (i + 1).ToString()].Text =
                    Transcode.CodeToSfjbnt(dtField.Rows[i][14].ToString());
                doc.Range.Bookmarks["d" + (i + 1).ToString()].Text = EditSz(dtField.Rows[i][4].ToString());
                doc.Range.Bookmarks["n" + (i + 1).ToString()].Text = EditSz(dtField.Rows[i][5].ToString());
                doc.Range.Bookmarks["x" + (i + 1).ToString()].Text = EditSz(dtField.Rows[i][6].ToString());
                doc.Range.Bookmarks["b" + (i + 1).ToString()].Text = EditSz(dtField.Rows[i][7].ToString());
            }
            doc.Range.Bookmarks["htsum"].Text = htSum.ToString("f");
            doc.Range.Bookmarks["scsum"].Text = scSum.ToString("f");
            doc.Range.Bookmarks["dkzs"].Text = dtField.Rows.Count.ToString();
            doc.Save(savedFilePath, SaveFormat.Pdf);

        }
    }
}
