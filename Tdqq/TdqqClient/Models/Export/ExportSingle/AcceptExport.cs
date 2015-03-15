using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using TdqqClient.Services.Common;

namespace TdqqClient.Models.Export.ExportSingle
{
    class AcceptExport:ExportBase
    {
        public AcceptExport(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }

        public override void Export(string cbfmc, string cbfbm, string folderPath, string edgeFeature = "")
        {
            var saveFilePath = folderPath + @"\" + cbfbm.Substring(14) + "_" + cbfmc + "_06公示无异议声明书.pdf";
            ExportDoc(saveFilePath, cbfmc, cbfbm);
        }

        private void ExportDoc(string savedDocPathStatement, string cbfmc,string cbfbm)
        {

            var rowFbf = SelectFbfInfo();
            if (rowFbf == null) return;
            var fbfdz = rowFbf[6].ToString().Trim();
            var rowCbf = SelectCbfInfoByCbfbm(cbfbm);
            //导出word
            var templateDocPath = AppDomain.CurrentDomain.BaseDirectory + @"\template\无异议声明书.doc";
            Document exportWord = new Document(templateDocPath);
            exportWord.Range.Bookmarks["承包方名称"].Text = cbfmc;
            exportWord.Range.Bookmarks["承包方住所"].Text = fbfdz;
            exportWord.Range.Bookmarks["性别2"].Text = Transcode.CodeToSex(rowCbf[3].ToString().Trim());
            exportWord.Range.Bookmarks["身份证号2"].Text = rowCbf[5].ToString().Trim();
            exportWord.Range.Bookmarks["承包方住所4"].Text = fbfdz;

            TimeSpan timeSpan = new TimeSpan(-9, 0, 0, 0);


            exportWord.Range.Bookmarks["日期2"].Text = GetShrq(-2).Add(timeSpan).ToLongDateString();
            exportWord.Range.Bookmarks["日期3"].Text = GetShrq(-2).ToLongDateString();
            exportWord.Save(savedDocPathStatement, SaveFormat.Pdf);
        }
    }
}
