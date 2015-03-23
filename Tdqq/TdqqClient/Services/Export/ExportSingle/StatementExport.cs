using System;
using System.Text;
using Aspose.Words;
using TdqqClient.Services.Common;

namespace TdqqClient.Services.Export.ExportSingle
{
    class StatementExport:ExportBase
    {

        public StatementExport(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        { }

        public override void Export(string cbfmc, string cbfbm, string folderPath, string edgeFeature = "")
        {
            var saveFilePath = folderPath + @"\" + cbfbm.Substring(14) + "_" + cbfmc + "_05户主声明书.pdf";
            ExportDoc(saveFilePath, cbfmc, cbfbm);
        }
        private void ExportDoc(string savedDocPathStatement, string cbfmc, string cbfbm)
        {
            var fbf = Fbf();
            if (fbf == null) return;
            var fbfdz =fbf.Fbfdz;
            var cbfjtcys = Cbfjtcys(cbfbm);
            var xb = Transcode.CodeToSex(cbfjtcys[0].Cyxb);
            var sfzhm = cbfjtcys[0].Cyzjhm;
            var stringBuilder = new StringBuilder();
            stringBuilder.Append(cbfmc);
            for (int i = 1; i < cbfjtcys.Count; i++)
            {
                stringBuilder.Append("、" + cbfjtcys[i].Cyxm);
            }
            stringBuilder.Append("共计" + cbfjtcys.Count + "人");
            var templateDocPath = AppDomain.CurrentDomain.BaseDirectory + @"\template\农户声明书.doc";
            Document exportWord = new Document(templateDocPath);
            exportWord.Range.Bookmarks["承包方名称2"].Text = cbfmc;
            exportWord.Range.Bookmarks["承包方住所2"].Text = fbfdz;
            exportWord.Range.Bookmarks["性别1"].Text = xb;
            exportWord.Range.Bookmarks["身份证号1"].Text = sfzhm;
            exportWord.Range.Bookmarks["家庭成员信息"].Text = stringBuilder.ToString();
            var dcsh = DcSh();
            exportWord.Range.Bookmarks["日期1"].Text = dcsh.Gsshrq.Add(new TimeSpan(-2,0,0,0)).ToLongDateString();
            exportWord.Save(savedDocPathStatement, SaveFormat.Pdf);
        }
    }
}
