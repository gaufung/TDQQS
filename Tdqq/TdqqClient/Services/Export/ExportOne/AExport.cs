using System;
using System.Windows.Forms;
using Aspose.Words;
using TdqqClient.Models;
using TdqqClient.Services.Common;

namespace TdqqClient.Services.Export.ExportOne
{
    /// <summary>
    /// 导出A表，即发包方调查表
    /// </summary>
    class AExport:ExportBase,IExport
    {
        public AExport(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }
        public void Export()
        {
            var fbf = Fbf();
            if (fbf==null)
            {
                MessageBox.Show(null, "请导入发包方信息", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            var dialogHelper = new DialogHelper("pdf");
            var savedPath = dialogHelper.SaveFile("保存发包方信息");
            if (string.IsNullOrEmpty(savedPath)) return;
            if (Export(fbf, savedPath))
            {
                MessageBox.Show(null, "发包方信息导出成功", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(null, "发包方信息导出失败", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool Export(FbfModel fbf, string saveFilePath)
        {
            try
            {
                var templatePath = AppDomain.CurrentDomain.BaseDirectory + @"\template\发包方调查表.doc";
                const string missInfo = "/";
                var doc = new Document(templatePath);
                var fbfmc = fbf.Fbfmc;
                doc.Range.Bookmarks["fbfmc"].Text = fbfmc.IndexOf("县") == -1
                    ? missInfo
                    : fbfmc.Substring(fbfmc.IndexOf("县") + 1);
                doc.Range.Bookmarks["fbfbm"].Text = string.IsNullOrEmpty(fbf.Fbfbm)?missInfo:fbf.Fbfbm;
                doc.Range.Bookmarks["fbffzrxm"].Text = string.IsNullOrEmpty(fbf.Fbffzrxm)? missInfo: fbf.Fbffzrxm;
                doc.Range.Bookmarks["lxdh"].Text = string.IsNullOrEmpty(fbf.Lxdh)? missInfo: fbf.Lxdh;
                doc.Range.Bookmarks["fbfdz"].Text = string.IsNullOrEmpty(fbf.Fbfdz)? missInfo: fbf.Fbfdz;
                doc.Range.Bookmarks["yzbm"].Text = string.IsNullOrEmpty(fbf.Yzbm)? missInfo: fbf.Fbfdz;
                if (!string.IsNullOrEmpty(fbf.Fzrzjlx)) doc.Range.Bookmarks["fbfzjlx"].Text = Transcode.Fbfzjlx(fbf.Fzrzjlx);
                doc.Range.Bookmarks["zjhm"].Text = string.IsNullOrEmpty(fbf.Fzrzjhm)? missInfo: fbf.Fzrzjhm;
                var dcsh = DcSh();
                if (dcsh!=null)
                {
                    doc.Range.Bookmarks["dcy"].Text = string.IsNullOrEmpty(dcsh.Cbfdcy) ? missInfo : dcsh.Cbfdcy;
                    doc.Range.Bookmarks["dcrq"].Text =dcsh.Cbfdcrq.ToLongDateString();
                    doc.Range.Bookmarks["shrq"].Text = dcsh.Cbfdcrq.Add(new TimeSpan(3, 0, 0, 0)).ToLongDateString();
                }                
                doc.Save(saveFilePath, SaveFormat.Pdf);
                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                return false;
            }   
        }
    }
}
