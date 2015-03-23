using System;
using Aspose.Words;

namespace TdqqClient.Services.Export.ExportSingle
{
    /// <summary>
    /// 导出档案局封面
    /// </summary>
    class CoverExport:ExportBase
    {

        public CoverExport(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        { }

        public override void Export(string cbfmc, string cbfbm, string folderPath, string edgeFeature = "")
        {
            var saveFilePath = folderPath + @"\" + cbfbm.Substring(14) + "_" + cbfmc + "_00档案局目录.pdf";
            ExportDoc(saveFilePath, cbfmc, cbfbm);
        }
        /// <summary>
        /// 导出某一个封面
        /// </summary>
        /// <param name="saveFilePath">路径</param>
        /// <param name="cbfmc">承包方名称</param>
        /// <param name="cbfbm">承包方编码</param>
        private void ExportDoc(string saveFilePath, string cbfmc, string cbfbm)
        {
            
            var dtFbf = Fbf();
            if (dtFbf == null) return;
            var fbfdz = Fbf().Fbfdz;
            var indexXian = fbfdz.IndexOf("县");
            var indexZhen = fbfdz.IndexOf("镇") == -1 ? fbfdz.IndexOf("乡") : fbfdz.IndexOf("镇");
            var indexCun = fbfdz.IndexOf("村");
            var zhen = fbfdz.Substring(indexXian + 1, indexZhen - indexXian - 1);
            var cun = fbfdz.Substring(indexZhen + 1, indexCun - indexZhen);
            var templatePath = AppDomain.CurrentDomain.BaseDirectory + @"\template\档案局目录.doc";
            var exportWord = new Document(templatePath);
            exportWord.Range.Bookmarks["cbfbm"].Text = cbfbm.Substring(14);
            exportWord.Range.Bookmarks["xiang"].Text = zhen;
            exportWord.Range.Bookmarks["cun"].Text = cun;
            exportWord.Range.Bookmarks["cbfmc"].Text = cbfmc;
            exportWord.Save(saveFilePath, SaveFormat.Pdf);
        }
    }
}
