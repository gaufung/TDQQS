using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using TdqqClient.Services.Common;

namespace TdqqClient.Models.Export.ExportSingle
{
    class StatementExport:ExportBase
    {
        public StatementExport(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }

        public override void Export(string cbfmc, string cbfbm, string folderPath, string edgeFeature = "")
        {
            var saveFilePath = folderPath + @"\" + cbfbm.Substring(14) + "_" + cbfmc + "_05户主声明书.pdf";
            ExportDoc(saveFilePath, cbfmc, cbfbm);
        }
        private void ExportDoc(string savedDocPathStatement, string cbfmc,string cbfbm)
        {
            var rowFbf = SelectFbfInfo();
            if (rowFbf == null) return;
            var fbfdz = rowFbf[6].ToString();
            var dtcbf = SelectCbf_JtcyByCbfbm(cbfbm);

            var rowCbf = SelectCbfInfoByCbfbm(cbfbm);
            var xb = Transcode.CodeToSex(dtcbf.Rows[0][2].ToString().Trim());
            var sfzhm = dtcbf.Rows[0][4].ToString().Trim();
            var stringBuilder = new StringBuilder();
            stringBuilder.Append(cbfmc);
            for (int i = 1; i < dtcbf.Rows.Count; i++)
            {
                stringBuilder.Append("、" + dtcbf.Rows[i][3].ToString().Trim());
            }
            stringBuilder.Append("共计" + dtcbf.Rows.Count.ToString() + "人");
            var templateDocPath = AppDomain.CurrentDomain.BaseDirectory + @"\template\农户声明书.doc";
            Document exportWord = new Document(templateDocPath);
            exportWord.Range.Bookmarks["承包方名称2"].Text = cbfmc;
            exportWord.Range.Bookmarks["承包方住所2"].Text = fbfdz;
            exportWord.Range.Bookmarks["性别1"].Text = xb;
            exportWord.Range.Bookmarks["身份证号1"].Text = sfzhm;
            exportWord.Range.Bookmarks["家庭成员信息"].Text = stringBuilder.ToString();
            exportWord.Range.Bookmarks["日期1"].Text = GetShrq(-2).ToLongDateString();
            exportWord.Save(savedDocPathStatement, SaveFormat.Pdf);
        }
    }
}
