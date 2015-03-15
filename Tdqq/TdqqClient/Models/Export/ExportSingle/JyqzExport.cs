using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using TdqqClient.Services.Common;

namespace TdqqClient.Models.Export.ExportSingle
{
    /// <summary>
    /// 导出单户的经营权证
    /// </summary>
    class JyqzExport:ExportBase
    {
        public JyqzExport(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }

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
            dir.CreateSubdirectory(cbfbm.Substring(14)+cbfmc);
            var subFolder = Path.Combine(folderPath, cbfbm.Substring(14, 4) + cbfmc);
            var saveDocPath = subFolder + @"\" + cbfbm.Substring(14) + "_" + cbfmc + "_经营权证.doc";
            var saveAppendixPath = subFolder + @"\" + cbfbm.Substring(14) + "_" + cbfmc + "_承包方信息附录.pdf";
            ExportCertifiCation(saveDocPath,cbfbm);
            ExportAppendix(saveAppendixPath,cbfbm);
        }
        /// <summary>
        /// 导出证书信息
        /// </summary>
        /// <param name="toSaveFilePath">保存文件的地址</param>
        /// <param name="cbfbm">承包方编码</param>
        private void ExportCertifiCation(string toSaveFilePath, string cbfbm)
        {
            /*  
             * CBFBM,CBFLX,CBFMC,CYXB,CBFZJLX,CBFZJHM,CBFDZ,YZBM,LXDH,CBFCYSL,CBFDCRQ,CBFDCY,CBFDCJS,GSJS,GSJSR,GSSHRQ
             * 0            1                2      3         4               5               6          7        8         9                10                  11           12         13     14          15
             */
            var rowCbf = SelectCbfInfoByCbfbm(cbfbm);
            if (rowCbf == null) return;
            var dtFields = SelectFieldsByCbfbm(cbfbm);
            var cbfmc = rowCbf[2].ToString();
            var cbfzjhm = rowCbf[5].ToString();
            var yzbm = rowCbf[7].ToString();
            var lxdh = rowCbf[8].ToString();
            DateTime dateTime = GetShrq(8);
            var year = dateTime.Year.ToString();
            var month = dateTime.Month.ToString();
            var day = dateTime.Day.ToString();
            //获取发包方信息
            /*
             * FBFBM,FBFMC,FBFFZRXM,FZRZJLX,FZRZJHM,LXDH,FBFDZ,YZBM,FBFDCY,FBFDCRQ,FBFDCJS
             */
            var rowFbf = SelectFbfInfo();
            var fbfmc = rowFbf[1].ToString().Substring(rowFbf[1].ToString().IndexOf("县") + 1);
            string xian, zhen, cun;
            xian = zhen = cun = string.Empty;
            SpiltFbfdzToXianXiangCun(rowFbf[6].ToString(), ref xian, ref zhen, ref cun);
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
            double sumScmj = 0.0;
            int fieldCount = dtFields.Rows.Count;
            /*
             * CBFMC,DKMC,YHTMJ,DKBM,DKDZ,DKNZ,DKXZ,DKBZ,DKBZXX,ZJRXM,DKLB,TDLYLX,DLDJ,TDYT,SFJBNT,CBJYQQDFS,HTMJ,SCMJ
             *      0           1           2          3        4       5           6       7       8               9       10          11          12       13       14            15              16      17
             */
            for (int i = 0; i < fieldCount; i++)
            {
                if (string.IsNullOrEmpty(dtFields.Rows[i][17].ToString().Trim()))
                {
                    sumScmj += 0.0;
                }
                else
                {
                    sumScmj += Convert.ToDouble(double.Parse(dtFields.Rows[i][17].ToString().Trim()).ToString("f"));
                }
            }
            exportWord.Range.Bookmarks["实测面积"].Text = sumScmj.ToString("f");
            exportWord.Range.Bookmarks["地块数"].Text = fieldCount.ToString();
            for (int i = 0; i < fieldCount; i++)
            {
                exportWord.Range.Bookmarks["地块名称" + i.ToString()].Text = dtFields.Rows[i][1].ToString();
                exportWord.Range.Bookmarks["地块编码" + i.ToString()].Text = dtFields.Rows[i][3].ToString();
                if (string.IsNullOrEmpty(dtFields.Rows[i][17].ToString().Trim()))
                {
                    exportWord.Range.Bookmarks["实测面积" + i.ToString()].Text = 0.0.ToString("f");
                }
                else
                {
                    exportWord.Range.Bookmarks["实测面积" + i.ToString()].Text = Convert.ToDouble(dtFields.Rows[i][17].ToString().Trim()).ToString("f");
                }
                exportWord.Range.Bookmarks["基本农田" + i.ToString()].Text = Transcode.CodeToSfjbnt(dtFields.Rows[i][14].ToString());
                exportWord.Range.Bookmarks["东" + i.ToString()].Text = EditSz(dtFields.Rows[i][4].ToString()) + "\r\n";
                exportWord.Range.Bookmarks["西" + i.ToString()].Text = EditSz(dtFields.Rows[i][5].ToString()) + "\r\n";
                exportWord.Range.Bookmarks["南" + i.ToString()].Text = EditSz(dtFields.Rows[i][6].ToString()) + "\r\n";
                exportWord.Range.Bookmarks["北" + i.ToString()].Text = EditSz(dtFields.Rows[i][7].ToString());
            }
            for (int i = fieldCount; i < 17; i++)
            {
                exportWord.Range.Bookmarks["基本农田" + i.ToString()].Text = " ";
                exportWord.Range.Bookmarks["地块名称" + i.ToString()].Text = " ";
                exportWord.Range.Bookmarks["地块编码" + i.ToString()].Text = " ";
                exportWord.Range.Bookmarks["实测面积" + i.ToString()].Text = " ";
                exportWord.Range.Bookmarks["东" + i.ToString()].Text = " ";
                exportWord.Range.Bookmarks["西" + i.ToString()].Text = " ";
                exportWord.Range.Bookmarks["南" + i.ToString()].Text = " ";
                exportWord.Range.Bookmarks["北" + i.ToString()].Text = " ";
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
            /*
            * FBFBM,FBFMC,FBFFZRXM,FZRZJLX,FZRZJHM,LXDH,FBFDZ,YZBM,FBFDCY,FBFDCRQ,FBFDCJS
            */
            var rowFbf = SelectFbfInfo();
            /*
            * CBFBM,CBFLX,CBFMC,CYXB,CBFZJLX,CBFZJHM,CBFDZ,YZBM,LXDH,CBFCYSL,CBFDCRQ,CBFDCY,CBFDCJS,GSJS,GSJSR,GSSHRQ
            * 0            1                2      3         4               5               6          7        8         9                10                  11           12         13     14          15
            */
            var rowCbf = SelectCbfInfoByCbfbm(cbfbm);
            var dtCbfJtcy = SelectCbf_JtcyByCbfbm(cbfbm);
            var fileDocSourth = AppDomain.CurrentDomain.BaseDirectory + @"\template\证书附录.doc";
            Document doc = new Document(fileDocSourth);
            DocumentBuilder documentBuilder = new DocumentBuilder(doc);
            //书签处插入
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
            for (int i = 0; i < dtCbfJtcy.Rows.Count; i++)
            {
                var name = dtCbfJtcy.Rows[i][1].ToString();
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
                        documentBuilder.Write(dtCbfJtcy.Rows[i][6].ToString());
                    }
                }
                documentBuilder.EndRow();
            }
            doc.Save(toSaveFilePath, SaveFormat.Pdf);
        }  
    }
}
