using System;
using System.Collections;
using System.Threading;
using System.Windows.Forms;
using TdqqClient.Services.Common;
using TdqqClient.ViewModels;
using TdqqClient.Views;

namespace TdqqClient.Services.Export.ExportTotal
{
    /// <summary>
    /// 导出整体
    /// </summary>
    class ExportTotal:ExportBase,IExport
    {
        private ExportBase _exportBase;
        private readonly bool _isNeedCunEdge;
        private readonly string _exportInfo;
        public ExportTotal(ExportBase export, bool isNeedCunEdge,string exportInfo, string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {
            _exportBase = export;
            _isNeedCunEdge = isNeedCunEdge;
            _exportInfo = exportInfo;
        }

        public void Export()
        {
            string cunEdge = string.Empty;
            if (_isNeedCunEdge)
            {
                SelectFeatureViewModel selectFeatureVm = new SelectFeatureViewModel(PersonDatabase);
                selectFeatureVm.Caption = "请选择村界要素类";
                SelectFeatureWindow selectFeautreV = new SelectFeatureWindow(selectFeatureVm);
                selectFeautreV.ShowDialog();
                cunEdge = selectFeatureVm.SelectFeature;
                if (string.IsNullOrEmpty(cunEdge)) return;
            }
            var dialogHelper = new DialogHelper();
            var folderPath = dialogHelper.OpenFolderDialog(true);
            if (string.IsNullOrEmpty(folderPath)) return;
            if (Export(folderPath,cunEdge))
            {
                MessageBox.Show(null, _exportInfo+"导出成功",
                    "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(null,
                    _exportInfo+"导出失败",
                    "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool Export(string folderPath,string cunEdge)
        {           
            Wait wait = new Wait();
            wait.SetWaitCaption("导出"+_exportInfo);
            Hashtable para = new Hashtable()
            {
                {"wait",wait},{"folderPath",folderPath},{"ret",false},{"cunEdge",cunEdge}
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
                var cbfs = Cbfs(true);
                var rowCount = cbfs.Count;
                for (int i = 0; i < rowCount; i++)
                {
                    wait.SetProgress(((double)i / (double)rowCount));
                    var cbfmc = cbfs[i].Cbfmc;
                    var cbfbm = cbfs[i].Cbfbm;
                    if (_isNeedCunEdge)
                    {
                        _exportBase.Export(cbfmc, cbfbm, folderPath, cunEdge);
                    }
                    else
                    {
                        _exportBase.Export(cbfmc, cbfbm, folderPath);
                    }
                    
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
