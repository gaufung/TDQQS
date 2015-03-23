using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using TdqqClient.Models;
using TdqqClient.Services.Common;

namespace TdqqClient.Services.Export.ExportSingle
{
    class ContractExport:ExportBase
    {

        public ContractExport(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        { }

        public override void Export(string cbfmc, string cbfbm, string folderPath, string edgeFeature = "")
        {
            var saveFilePath = folderPath + @"\" + cbfbm.Substring(14) + "_" + cbfmc + "_03承包经营权合同书.pdf";
            ExportDoc(saveFilePath, cbfmc, cbfbm);
        }

        private void ExportDoc(string saveFilePath, string cbfmc, string cbfbm)
        {
           
            var rowFbf = Fbf();
            if (rowFbf == null) return;
            var rowCbf = Cbf(cbfbm);
            if (rowCbf == null) return;
            ExportDoc(saveFilePath, rowFbf, rowCbf);
        }

        private void ExportDoc(string saveFilePath, FbfModel fbf, CbfModel cbf)
        {
            
            var fields =Fields(cbf.Cbfbm);
            if (fields == null) return;
            var dtCyxx = Cbfjtcys(cbf.Cbfbm);
            if (dtCyxx == null) return;
            if (fields.Count > 17)
            {
                System.Windows.Forms.MessageBox.Show(cbf.Cbfbm + "该农户的地块数超过17块");
                return;
            }
            var docTemplatePath = AppDomain.CurrentDomain.BaseDirectory + @"\template\合同书\合同" + fields.Count + ".doc";
            var scmjSum = AreaSum(fields, true);
            Document doc = new Document(docTemplatePath);
            doc.Range.Bookmarks["合同编号"].Text = cbf.Cbfbm + "J";
            doc.Range.Bookmarks["id"].Text = cbf.Cbfzjhm;
            doc.Range.Bookmarks["发包方名称"].Text = fbf.Fbfmc;
            doc.Range.Bookmarks["承包方名称1"].Text = cbf.Cbfmc;
            doc.Range.Bookmarks["承包方住所1"].Text =fbf.Fbfdz;
            doc.Range.Bookmarks["实测面积"].Text = scmjSum.ToString("F");
            for (int i = 0; i < fields.Count; i++)
            {
                doc.Range.Bookmarks["dkmc" + (i + 1)].Text = fields[i].Dkmc;
                doc.Range.Bookmarks["dkbm" + (i + 1)].Text = fields[i].Dkbm;
                doc.Range.Bookmarks["scmj" + (i + 1)].Text = fields[i].Scmj.ToString("F");
                doc.Range.Bookmarks["scmj" + (i + 1)].Text = fields[i].Htmj.ToString("F");
                doc.Range.Bookmarks["dz" + (i + 1)].Text = fields[i].Dkdz;
                doc.Range.Bookmarks["nz" + (i + 1)].Text = fields[i].Dknz;
                doc.Range.Bookmarks["xz" + (i + 1)].Text = fields[i].Dkxz;
                doc.Range.Bookmarks["bz" + (i + 1)].Text = fields[i].Dkbz;
                doc.Range.Bookmarks["sf" + (i + 1)].Text = Transcode.CodeToSfjbnt(fields[i].Sfjbnt);
            }
            doc.Range.Bookmarks["大写"].Text = ConvertNumberHelper.ConvertSum(scmjSum.ToString());
            doc.Range.Bookmarks["小写"].Text = scmjSum.ToString("f");
            doc.Range.Bookmarks["地块"].Text = fields.Count.ToString();
            var dcsh = DcSh();
            doc.Range.Bookmarks["rq1"].Text = dcsh.Gsshrq.Add(new TimeSpan(5, 0, 0, 0)).ToLongDateString();
            doc.Range.Bookmarks["rq2"].Text = dcsh.Gsshrq.Add(new TimeSpan(5, 0, 0, 0)).ToLongDateString();
            doc.Range.Bookmarks["rq3"].Text = dcsh.Gsshrq.Add(new TimeSpan(5, 0, 0, 0)).ToLongDateString();
            doc.Save(saveFilePath, SaveFormat.Pdf);
        }
        /// <summary>
        /// 获取总面积
        /// </summary>
        /// <param name="fields">某一户地块的集合</param>
        /// <param name="isScmj">是否为实测面积</param>
        /// <returns>总面积</returns>
        private double AreaSum(IEnumerable<FieldModel> fields, bool isScmj)
        {
            if (isScmj)
            {
                var query = from field in fields
                    select field.Scmj;
                return query.Sum();
            }
            else
            {
                var query = from field in fields
                            select field.Htmj;
                return query.Sum();
            }

        }
    }
}
