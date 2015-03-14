using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Aspose.Words;
using TdqqClient.Services.Common;

namespace TdqqClient.Models.Export.ExportOne
{
    /// <summary>
    /// 导出
    /// </summary>
    class AExport:ExportOne
    {
        public AExport(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }
        public override void Export(object parameter)
        {
            var fbfInfo = SelectFbfInfo();
            if (fbfInfo==null)
            {
                MessageBox.Show(null, "请导入发包方信息", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            DialogHelper dialogHelper=new DialogHelper("pdf");
            var savedPath = dialogHelper.SaveFile("保存发包方信息");
            if (string.IsNullOrEmpty(savedPath)) return;
            if (Export(fbfInfo, savedPath))
            {
                MessageBox.Show(null, "发包方信息导出成功", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(null, "发包方信息导出失败", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }
        /// <summary>
        /// 导出发包方信息
        /// </summary>
        /// <param name="rowInfo">发包方一行数据</param>
        /// <param name="saveFilePath">已经要保存的路径</param>
        /// <returns></returns>
        private bool Export(System.Data.DataRow rowInfo, string saveFilePath)
        {
            try
            {
                var templatePath = AppDomain.CurrentDomain.BaseDirectory + @"\template\发包方调查表.doc";
                const string missInfo = "/";
                var doc = new Document(templatePath);
                var fbfmc = rowInfo[1].ToString().Trim();
                doc.Range.Bookmarks["fbfmc"].Text = fbfmc.IndexOf("县") == -1
                    ? missInfo
                    : fbfmc.Substring(fbfmc.IndexOf("县") + 1);
                doc.Range.Bookmarks["fbfbm"].Text = string.IsNullOrEmpty(rowInfo[0].ToString().Trim())
                    ? missInfo
                    : rowInfo[0].ToString();
                doc.Range.Bookmarks["fbffzrxm"].Text = string.IsNullOrEmpty(rowInfo[2].ToString().Trim())
                    ? missInfo
                    : rowInfo[2].ToString().Trim();
                doc.Range.Bookmarks["lxdh"].Text = string.IsNullOrEmpty(rowInfo[5].ToString().Trim())
                    ? missInfo
                    : rowInfo[5].ToString().Trim();
                doc.Range.Bookmarks["fbfdz"].Text = string.IsNullOrEmpty(rowInfo[6].ToString().Trim())
                    ? missInfo
                    : rowInfo[6].ToString().Trim();
                doc.Range.Bookmarks["yzbm"].Text = string.IsNullOrEmpty(rowInfo[7].ToString().Trim())
                    ? missInfo
                    : rowInfo[7].ToString().Trim();
                if (!string.IsNullOrEmpty(Transcode.Fbfzjlx(rowInfo[3].ToString().Trim())))
                {
                    doc.Range.Bookmarks["fbfzjlx"].Text = Transcode.Fbfzjlx(rowInfo[3].ToString().Trim());
                }
                doc.Range.Bookmarks["zjhm"].Text = string.IsNullOrEmpty(rowInfo[4].ToString().Trim())
                    ? missInfo
                    : rowInfo[4].ToString().Trim();
                doc.Range.Bookmarks["dcy"].Text = string.IsNullOrEmpty(GetDcy()) ? missInfo : GetDcy();
                doc.Range.Bookmarks["dcrq"].Text = GetDcrq().ToLongDateString();
                doc.Range.Bookmarks["shrq"].Text = GetDcrq(3).ToLongDateString();
                doc.Save(saveFilePath, SaveFormat.Pdf);
                return true;   
            }
            catch (Exception)
            {
                return false;
            }
           
        }
    }
}
