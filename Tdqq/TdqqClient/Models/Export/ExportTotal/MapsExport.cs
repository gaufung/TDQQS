using System;
using System.Collections;
using System.Threading;
using System.Windows.Forms;
using TdqqClient.Models.Export.ExportSingle;
using TdqqClient.Services.Common;
using TdqqClient.ViewModels;
using TdqqClient.Views;

namespace TdqqClient.Models.Export.ExportTotal
{
    class MapsExport:ExportBase
    {
        public MapsExport(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }
        public void Export()
        {
            SelectFeatureViewModel selectFeatureVm=new SelectFeatureViewModel(PersonDatabase);
            selectFeatureVm.Caption = "请选择村界要素类";
            SelectFeatureWindow selectFeautreV=new SelectFeatureWindow(selectFeatureVm);
            selectFeautreV.ShowDialog();
            var cunEdge = selectFeatureVm.SelectFeature;
            if (string.IsNullOrEmpty(cunEdge)) return;
            var dialogHelper = new DialogHelper();
            var folderPath = dialogHelper.OpenFolderDialog(true);
            if (string.IsNullOrEmpty(folderPath)) return;
            if (Export(folderPath,cunEdge))
            {
                MessageBox.Show(null, "地块示意图导出成功",
                    "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(null,
                    "地块示意图导出失败",
                    "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool Export(string folderPath,string cunEdge)
        {
            Wait wait = new Wait();
            wait.SetWaitCaption("导出地块示意图");
            Hashtable para = new Hashtable()
            {
                {"wait",wait},{"folderPath",folderPath},{"cunEdge",cunEdge},{"ret",false}
            };
            Thread t = new Thread(new ParameterizedThreadStart(ExportF));
            t.Start(para);
            wait.ShowDialog();
            t.Abort();
            return (bool)para["ret"];
        }

        private void ExportF(object p)
        {
            var para = p as Hashtable;
            var wait = para["wait"] as Wait;
            var folderPath = para["folderPath"].ToString().Trim();
            var cunEdge = para["cunEdge"].ToString();
            try
            {
                var dt = SelectCbfbmOwnFields();
                var rowCount = dt.Rows.Count;
                ExportBase export = new MapExport(PersonDatabase, SelectFeature, BasicDatabase);
                for (int i = 0; i < rowCount; i++)
                {
                    wait.SetProgress(((double)i / (double)rowCount));
                    var cbfmc = dt.Rows[i][1].ToString().Trim();
                    var cbfbm = dt.Rows[i][0].ToString().Trim();
                    export.Export(cbfmc, cbfbm, folderPath, cunEdge);
                }
                para["ret"] = true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                para["ret"] = false;
            }
            finally
            {
                wait.CloseWait();
            }
        }
    }
}
