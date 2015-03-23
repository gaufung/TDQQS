using System;
using Aspose.Words;
using NPOI.HPSF;
using TdqqClient.Services.Common;

namespace TdqqClient.Services.Export.ExportSingle
{
    class CbfExport:ExportBase
    {
        public CbfExport(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }
        public override void Export(string cbfmc, string cbfbm, string folderPath, string edgeFeature = "")
        {
            var saveFilePath = folderPath + @"\" + cbfbm.Substring(14) + "_" + cbfmc + "_01承包方调查表.pdf";
            ExportDoc(saveFilePath, cbfmc, cbfbm);
        }
        private void ExportDoc(string saveFilePath, string cbfmc, string cbfbm)
        {
            var templatePath = AppDomain.CurrentDomain.BaseDirectory + @"\template\承包方调查表.doc";
            var cbfjtcys = Cbfjtcys(cbfbm);
            if (cbfjtcys == null||cbfjtcys.Count==0) return;
            var doc = new Document(templatePath);
            var documentBuilder = new DocumentBuilder(doc);
            doc.Range.Bookmarks["fbfbm"].Text = cbfbm.Substring(0, 14);
            doc.Range.Bookmarks["cbfbm"].Text = cbfbm;
            doc.Range.Bookmarks["cbfmc"].Text = cbfmc;
            var lxdh = string.IsNullOrEmpty(cbfjtcys[0].Lxdh) ?
                "/" : cbfjtcys[0].Lxdh;
            doc.Range.Bookmarks["lxdh"].Text = lxdh;
            doc.Range.Bookmarks["cbfdz"].Text = cbfjtcys[0].Cyszc;
            doc.Range.Bookmarks["yzbm"].Text = cbfjtcys[0].Yzbm;
            doc.Range.Bookmarks["zjhm"].Text = cbfjtcys[0].Cyzjhm;
            //开始插入承包方家庭成员信息
            int startRowIndex = 10;
            doc.Range.Bookmarks["cbfcysl"].Text = "共" + cbfjtcys.Count + "人";
            for (int i = 0; i < cbfjtcys.Count; i++)
            {
                documentBuilder.MoveToCell(0, startRowIndex + i, 0, 0);
                documentBuilder.Write(cbfjtcys[i].Cyxm);
                documentBuilder.MoveToCell(0, startRowIndex + i, 1, 0);
                documentBuilder.Write(Transcode.CodeToRelationship(cbfjtcys[i].Yhzgx));
                documentBuilder.MoveToCell(0, startRowIndex + i, 2, 0);
                var sfzh = string.IsNullOrEmpty(cbfjtcys[i].Cyzjhm)
                       ? "/"
                       : cbfjtcys[i].Cyzjhm;
                documentBuilder.Write(sfzh);
                var cybz = string.IsNullOrEmpty(cbfjtcys[i].Cybz)
                   ? "/"
                   : cbfjtcys[i].Cybz;
                documentBuilder.MoveToCell(0, startRowIndex + i, 3, 0);
                documentBuilder.Write(cybz);
            }
            var dcsh = DcSh();
            doc.Range.Bookmarks["dcy"].Text = dcsh.Cbfdcy;
            doc.Range.Bookmarks["rq1"].Text =dcsh.Cbfdcrq.ToLongDateString();
            doc.Range.Bookmarks["rq2"].Text = dcsh.Cbfdcrq.Add(new TimeSpan(3,0,0,0)).ToLongDateString();
            doc.Save(saveFilePath, SaveFormat.Pdf);
        }
    }
}
