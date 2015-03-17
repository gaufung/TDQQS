﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using TdqqClient.Models.Export.ExportSingle;
using TdqqClient.Services.Common;
using TdqqClient.Views;

namespace TdqqClient.Models.Export.ExportTotal
{
    class DksExport:ExportBase
    {
        public DksExport(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }
        public void Export()
        {
            var dialogHelper = new DialogHelper();
            var folderPath = dialogHelper.OpenFolderDialog(true);
            if (string.IsNullOrEmpty(folderPath)) return;
            if (Export(folderPath))
            {
                MessageBox.Show(null, "地块调查表导出成功",
                    "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(null,
                    "地块调查表导出失败",
                    "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool Export(string folderPath)
        {
            Wait wait = new Wait();
            wait.SetWaitCaption("导出地块调查表");
            Hashtable para = new Hashtable()
            {
                {"wait",wait},{"folderPath",folderPath},{"ret",false}
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
            try
            {
                var dt = SelectCbfbmOwnFields();
                var rowCount = dt.Rows.Count;
                ExportBase export = new DkExport(PersonDatabase, SelectFeature, BasicDatabase);
                for (int i = 0; i < rowCount; i++)
                {
                    wait.SetProgress(((double)i / (double)rowCount));
                    var cbfmc = dt.Rows[i][1].ToString().Trim();
                    var cbfbm = dt.Rows[i][0].ToString().Trim();
                    export.Export(cbfmc, cbfbm, folderPath);
                }
                //删除掉无效的地块
                //DeleteFiles(folderPath);
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