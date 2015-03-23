using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using TdqqClient.Services.Common;

namespace TdqqClient.Models.Export.ExportSingle
{
    class ContractExport:ExportBase
    {

        public ContractExport(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }

        public override void Export(string cbfmc, string cbfbm, string folderPath, string edgeFeature = "")
        {
            var saveFilePath = folderPath + @"\" + cbfbm.Substring(14) + "_" + cbfmc + "_03承包经营权合同书.pdf";
            ExportDoc(saveFilePath,cbfmc,cbfbm);
        }

        private void ExportDoc(string saveFilePath, string cbfmc, string cbfbm)
        {
            /*
            * select FBFBM,FBFMC,FBFFZRXM,FZRZJLX,FZRZJHM,LXDH,FBFDZ,YZBM,FBFDCY,FBFDCRQ,FBFDCJS
            */
            var rowFbf = SelectFbfInfo();
            if (rowFbf == null) return;
            /*
             * CBFBM,CBFLX,CBFMC,CYXB,CBFZJLX,CBFZJHM,CBFDZ,YZBM,LXDH,CBFCYSL,CBFDCRQ,CBFDCY,CBFDCJS,GSJS,GSJSR,GSSHRQ,GSSHR
             */
            var rowCbf = SelectCbfInfoByCbfbm(cbfbm);
            if (rowCbf == null) return;
            ExportDoc(saveFilePath, rowFbf, rowCbf);
        }

        private void ExportDoc(string saveFilePath, System.Data.DataRow rowFbf, System.Data.DataRow rowCbf)
        {
            var scmjSum = 0.0;
            /*
             * CBFMC,DKMC,YHTMJ,DKBM,DKDZ,DKNZ,DKXZ,DKBZ,DKBZXX,ZJRXM,DKLB,TDLYLX,DLDJ,TDYT,SFJBNT,CBJYQQDFS,HTMJ,SCMJ
             */
            var dtDk = SelectFieldsByCbfbm(rowCbf[0].ToString().Trim());
            if (dtDk == null) return;
            for (int i = 0; i < dtDk.Rows.Count; i++)
            {
                if (string.IsNullOrEmpty(dtDk.Rows[i][17].ToString()))
                {
                    scmjSum += 0.0;
                }
                else
                {
                    scmjSum += Convert.ToDouble(double.Parse(dtDk.Rows[i][17].ToString().Trim()).ToString("f"));
                }
            }
            var dtCyxx = SelectCbf_JtcyByCbfbm(rowCbf[0].ToString().Trim());
            if (dtCyxx == null) return;
            if (dtDk == null || dtDk.Rows.Count > 17)
            {
                System.Windows.Forms.MessageBox.Show(rowCbf[4].ToString() + "该农户的地块数超过17块");
                return;
            }
            var docTemplatePath = AppDomain.CurrentDomain.BaseDirectory + @"\template\合同书\合同" + dtDk.Rows.Count.ToString() + ".doc";
            Document doc = new Document(docTemplatePath);
            doc.Range.Bookmarks["合同编号"].Text = rowCbf[0].ToString().Trim() + "J";
            doc.Range.Bookmarks["id"].Text = rowCbf[5].ToString();
            doc.Range.Bookmarks["发包方名称"].Text = rowFbf[1].ToString();
            doc.Range.Bookmarks["承包方名称1"].Text = rowCbf[2].ToString();
            doc.Range.Bookmarks["承包方住所1"].Text = rowFbf[6].ToString();
            doc.Range.Bookmarks["实测面积"].Text = scmjSum.ToString("f");
            for (int i = 0; i < dtDk.Rows.Count; i++)
            {
                doc.Range.Bookmarks["dkmc" + (i + 1).ToString()].Text = dtDk.Rows[i][1].ToString();
                doc.Range.Bookmarks["dkbm" + (i + 1).ToString()].Text = dtDk.Rows[i][3].ToString();
                if (string.IsNullOrEmpty(dtDk.Rows[i][17].ToString()))
                {
                    doc.Range.Bookmarks["scmj" + (i + 1).ToString()].Text = 0.0.ToString("f");
                }
                else
                {
                    doc.Range.Bookmarks["scmj" + (i + 1).ToString()].Text = Convert.ToDouble(dtDk.Rows[i][17].ToString()).ToString("f");
                }
                doc.Range.Bookmarks["dz" + (i + 1).ToString()].Text = dtDk.Rows[i][4].ToString();
                doc.Range.Bookmarks["nz" + (i + 1).ToString()].Text = dtDk.Rows[i][5].ToString();
                doc.Range.Bookmarks["xz" + (i + 1).ToString()].Text = dtDk.Rows[i][6].ToString();
                doc.Range.Bookmarks["bz" + (i + 1).ToString()].Text = dtDk.Rows[i][7].ToString();
                doc.Range.Bookmarks["sf" + (i + 1).ToString()].Text = Transcode.CodeToSfjbnt(dtDk.Rows[i][14].ToString());
            }
            scmjSum = Convert.ToDouble(scmjSum.ToString("f"));
            doc.Range.Bookmarks["大写"].Text = ConvertNumberHelper.ConvertSum(scmjSum.ToString());
            doc.Range.Bookmarks["小写"].Text = scmjSum.ToString("f");
            doc.Range.Bookmarks["地块"].Text = dtDk.Rows.Count.ToString();
            doc.Range.Bookmarks["rq1"].Text = GetShrq(5).ToLongDateString();
            doc.Range.Bookmarks["rq2"].Text = GetShrq(5).ToLongDateString();
            doc.Range.Bookmarks["rq3"].Text = GetShrq(5).ToLongDateString();
            doc.Save(saveFilePath, SaveFormat.Pdf);   
        }
    }
}
