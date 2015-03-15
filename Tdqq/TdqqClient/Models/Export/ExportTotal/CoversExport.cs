using System;
using System.Collections;
using System.Threading;
using System.Windows.Forms;
using TdqqClient.Models.Export.ExportSingle;
using TdqqClient.Services.Common;
using TdqqClient.Views;

namespace TdqqClient.Models.Export.ExportTotal
{
    /// <summary>
    /// 导出整个村的封面
    /// </summary>
    class CoversExport:ExportBase
    {
        public CoversExport(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }

        public void Export()
        {
            var dialogHelper=new DialogHelper();
            var folderPath = dialogHelper.OpenFolderDialog(true);
            if (string.IsNullOrEmpty(folderPath)) return;
            if (Export(folderPath))
            {
                MessageBox.Show(null, "档案局目录导出成功",
                    "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(null, "档案局目录导出失败", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool Export(string folderPath)
        {
            var wait=new Wait();
            wait.SetWaitCaption("导出档案局目录");
            var para=new Hashtable()
            {
                {"wait",wait},{"folderPath",folderPath},{"ret",false}
            };
            var t=new Thread(new ParameterizedThreadStart(ExportF));
            t.Start(para);
            wait.ShowDialog();
            t.Abort();
            return (bool) para["ret"];
        }

        private void ExportF(object p)
        {
            Hashtable para = p as Hashtable;
            Wait wait = para["wait"] as Wait;
            var folderPath = para["folderPath"].ToString();
            try
            {
                var dt = SelectCbfbmOwnFields();
                var rowCount = dt.Rows.Count;
                ExportBase export=new CoverExport(PersonDatabase,SelectFeature,BasicDatabase);
                for (int i = 0; i < rowCount; i++)
                {
                    wait.SetProgress(((double)i / (double)rowCount));
                    var cbfmc = dt.Rows[i][1].ToString().Trim();
                    var cbfbm = dt.Rows[i][0].ToString().Trim();
                    export.Export(cbfmc, cbfbm, folderPath);
                }
                para["ret"] = true;
            }
            catch (Exception)
            {
                para["ret"] = false;
            }
            finally
            {
                wait.CloseWait();
            }
        }
    }
}
