using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;

namespace TdqqClient.Models.Export.ExportSingle
{
    class CoverExport:ExportBase
    {
        public CoverExport(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }

        public override void Export(string cbfmc, string cbfbm, string folderPath, string edgeFeature = "")
        {
            var saveFilePath = folderPath + @"\" + cbfbm.Substring(14) + "_" + cbfmc + "_00档案局目录.pdf";
            ExportDoc(saveFilePath,cbfmc,cbfbm);
        }

        private void ExportDoc(string savedDocPathStatement, string cbfmc,string cbfbm)
        {
            /*   "select FBFBM,FBFMC,FBFFZRXM,FZRZJLX,FZRZJHM,LXDH,FBFDZ,YZBM,FBFDCY,FBFDCRQ,FBFDCJS from FBF");*/
            var dtFbf = SelectFbfInfo();
            if (dtFbf == null) return;
            var fbfdz = dtFbf[6].ToString().Trim();
            var indexXian = fbfdz.IndexOf("县");
            var indexZhen = fbfdz.IndexOf("镇") == -1 ? fbfdz.IndexOf("乡") : fbfdz.IndexOf("镇");
            var indexCun = fbfdz.IndexOf("村");
            var zhen = fbfdz.Substring(indexXian + 1, indexZhen - indexXian - 1);
            var cun = fbfdz.Substring(indexZhen + 1, indexCun - indexZhen);
            var templatePath = AppDomain.CurrentDomain.BaseDirectory + @"\template\档案局目录.doc";
            Document exportWord = new Document(templatePath);
            exportWord.Range.Bookmarks["cbfbm"].Text = cbfbm.Substring(14);
            exportWord.Range.Bookmarks["xiang"].Text = zhen;
            exportWord.Range.Bookmarks["cun"].Text = cun;
            exportWord.Range.Bookmarks["cbfmc"].Text = cbfmc;
            exportWord.Save(savedDocPathStatement, SaveFormat.Pdf);
        }
    }
}
