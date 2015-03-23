using System;
using Aspose.Words;
using TdqqClient.Services.Common;

namespace TdqqClient.Services.Export.ExportSingle
{
    /// <summary>
    /// 导出某一户的无异议声明书
    /// </summary>
    class AcceptExport:ExportBase
    {
        public AcceptExport(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }
        /// <summary>
        /// 导出无异议
        /// </summary>
        /// <param name="cbfmc"></param>
        /// <param name="cbfbm"></param>
        /// <param name="folderPath"></param>
        /// <param name="edgeFeature"></param>
        public override  void Export(string cbfmc, string cbfbm, string folderPath, string edgeFeature = "")
        {
            var saveFilePath = folderPath + @"\" + cbfbm.Substring(14) + "_" + cbfmc + "_06公示无异议声明书.pdf";
            ExportDoc(saveFilePath, cbfmc, cbfbm);
        }
        /// <summary>
        ///导出文档
        /// </summary>
        /// <param name="saveFilePath">文件地址</param>
        /// <param name="cbfmc">承包方名称</param>
        /// <param name="cbfbm">承包方编码</param>
        private void ExportDoc(string saveFilePath, string cbfmc, string cbfbm)
        {
            var fbf = Fbf();
            if (fbf == null) return;
            var fbfdz = fbf.Fbfdz;
            var cbf = Cbf(cbfbm);
            //导出word
            var templateDocPath = AppDomain.CurrentDomain.BaseDirectory + @"\template\无异议声明书.doc";
            var exportWord = new Document(templateDocPath);
            exportWord.Range.Bookmarks["承包方名称"].Text = cbfmc;
            exportWord.Range.Bookmarks["承包方住所"].Text = fbfdz;
            exportWord.Range.Bookmarks["性别2"].Text = Transcode.CodeToSex(cbf.Cyxb);
            exportWord.Range.Bookmarks["身份证号2"].Text = cbf.Cbfzjhm;
            exportWord.Range.Bookmarks["承包方住所4"].Text = fbfdz;
            var dcsh = DcSh();
            exportWord.Range.Bookmarks["日期2"].Text = dcsh.Gsshrq.Add(new TimeSpan(-2,0,0,0)).Add(new TimeSpan(-9, 0, 0, 0)).ToLongDateString();
            exportWord.Range.Bookmarks["日期3"].Text = dcsh.Gsshrq.Add(new TimeSpan(-2, 0, 0, 0)).ToLongDateString();
            exportWord.Save(saveFilePath, SaveFormat.Pdf);
        }
    }
}
