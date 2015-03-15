using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using TdqqClient.Services.Common;

namespace TdqqClient.Models.Export.ExportSingle
{
    class CbfExport:ExportBase
    {
        public CbfExport(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }

        public override void Export(string cbfmc, string cbfbm, string folderPath, string edgeFeature = "")
        {
            var saveFilePath = folderPath + @"\" + cbfbm.Substring(14) + "_" + cbfmc + "_01承包方调查表.pdf";
            ExportDoc(saveFilePath,cbfmc,cbfbm);
        }
        private void ExportDoc(string saveFilePath, string cbfmc,string cbfbm)
        {
            var templatePath = AppDomain.CurrentDomain.BaseDirectory + @"\template\承包方调查表.doc";   
            var dt = SelectCbf_JtcyByCbfbm(cbfbm);
            /*
             *   "select CBFBM,CBFMC,CYXB,CYXM,CYZJHM,CYZJLX,CYBZ,CYBZ,YHZGX,CYSZC,YZBM,LXDH,SFGYR" +
                  " from CBF_JTCY where trim(CBFBM) = '{0}' order by YHZGX", cbfbm);
             */
            if (dt == null) return;
            var doc = new Document(templatePath);
            var documentBuilder = new DocumentBuilder(doc);
            doc.Range.Bookmarks["fbfbm"].Text = cbfbm.Substring(0, 14);
            doc.Range.Bookmarks["cbfbm"].Text = cbfbm;
            doc.Range.Bookmarks["cbfmc"].Text = cbfmc;
            var lxdh = string.IsNullOrEmpty(dt.Rows[0][11].ToString().Trim()) ? 
                "/" : dt.Rows[0][11].ToString().Trim();
            doc.Range.Bookmarks["lxdh"].Text = lxdh;
            doc.Range.Bookmarks["cbfdz"].Text = dt.Rows[0][9].ToString().Trim();
            doc.Range.Bookmarks["yzbm"].Text = dt.Rows[0][10].ToString().Trim();
            doc.Range.Bookmarks["zjhm"].Text = dt.Rows[0][4].ToString().Trim();
            //开始插入承包方家庭成员信息
            int startRowIndex = 10;
            doc.Range.Bookmarks["cbfcysl"].Text = "共" + dt.Rows.Count + "人";
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                documentBuilder.MoveToCell(0, startRowIndex + i, 0, 0);
                documentBuilder.Write(dt.Rows[i][3].ToString());
                documentBuilder.MoveToCell(0, startRowIndex + i, 1, 0);
                documentBuilder.Write(Transcode.CodeToRelationship(dt.Rows[i][8].ToString()));
                documentBuilder.MoveToCell(0, startRowIndex + i, 2, 0);
                var sfzh = string.IsNullOrEmpty(dt.Rows[i][4].ToString().Trim())
                       ? "/"
                       : dt.Rows[i][4].ToString().Trim();
                documentBuilder.Write(sfzh);
                var cybz = string.IsNullOrEmpty(dt.Rows[i][6].ToString().Trim())
                   ? "/"
                   : dt.Rows[i][6].ToString().Trim();
                documentBuilder.MoveToCell(0, startRowIndex + i, 3, 0);
                documentBuilder.Write(cybz);
            }
            doc.Range.Bookmarks["dcy"].Text = GetDcy();
            doc.Range.Bookmarks["rq1"].Text = GetDcrq().ToLongDateString();
            doc.Range.Bookmarks["rq2"].Text = GetDcrq(3).ToLongDateString();
            doc.Save(saveFilePath, SaveFormat.Pdf);
        }
    }
}
