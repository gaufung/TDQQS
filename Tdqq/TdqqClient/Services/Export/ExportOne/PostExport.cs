using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Aspose.Pdf.InteractiveFeatures.Forms;
using Aspose.Words;
using TdqqClient.Services.Common;
using TdqqClient.Services.Database;

namespace TdqqClient.Services.Export.ExportOne
{
    class PostExport:ExportBase,IExport
    {
        public PostExport(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }
        public void Export()
        {
            var fbfdz = GetFbfdz();
            if (string.IsNullOrEmpty(fbfdz))
            {
                MessageBox.Show(null, "发包方地址错误",
                    "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            var area = GetArea();
            if (area == null)
            {
                MessageBox.Show(null, "实测面积出错",
                    "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            var dialogHelper = new DialogHelper();
            var folderPath = dialogHelper.OpenFolderDialog(true);
            if (string.IsNullOrEmpty(folderPath)) return;
            if (Export(folderPath))
            {
                MessageBox.Show(null, "公示公告导出成功",
                    "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(null, "公示公告导出失败",
                    "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool Export(string folderPath)
        {
            var fbfdz = GetFbfdz();
            var name = SplitNameFromFbfdz(fbfdz);
            var countryPost = folderPath + @"\" + name + "村组公示公告.doc";
            var departmentPost = folderPath + @"\" + name + "部门审核公告.doc";
            int cbfCount = Cbfs(true).Count;
            int fieldCount = GetFieldCount();
            double area = (double)GetArea();
            bool ret = true;
            ret &= ExportCountry(countryPost, fbfdz, cbfCount, fieldCount, area);
            ret &= ExportDepartment(departmentPost, fbfdz, cbfCount, fieldCount, area);
            return ret;
        }

        private bool ExportCountry(string targetFilePath, string fbfdz, int cbfCount, int fieldCount, double area)
        {
            bool flag;
            try
            {
                var templatePath = AppDomain.CurrentDomain.BaseDirectory + @"\template\村组公示公告.doc";
                var name = SplitNameFromFbfdz(fbfdz);
                Document exportWord = new Document(templatePath);
                exportWord.Range.Bookmarks["户数"].Text = cbfCount.ToString();
                exportWord.Range.Bookmarks["宗地"].Text = fieldCount.ToString();
                exportWord.Range.Bookmarks["面积"].Text = area.ToString("f1");
                exportWord.Range.Bookmarks["发包方地址2"].Text = fbfdz;
                exportWord.Range.Bookmarks["发包方地址3"].Text = fbfdz;
                exportWord.Range.Bookmarks["发包方地址1"].Text = fbfdz;
                exportWord.Range.Bookmarks["村委会"].Text = name + "村民委员会";
                exportWord.Range.Bookmarks["发包方名称"].Text = fbfdz + "村民委员会";
                exportWord.Save(targetFilePath);
                flag = true;
            }
            catch (Exception)
            {
                flag = false;
            }
            return flag;

        }

        private bool ExportDepartment(string targetFilePath, string fbfdz, int cbfCount, int fieldCount, double area)
        {
            bool flag;
            try
            {
                var templatePath = AppDomain.CurrentDomain.BaseDirectory + @"\template\部门审核公告.doc";
                var name = SplitNameFromFbfdz(fbfdz);
                Document exportWord = new Document(templatePath);
                exportWord.Range.Bookmarks["发包方地址1"].Text = fbfdz;
                exportWord.Range.Bookmarks["村庄"].Text = name;
                exportWord.Range.Bookmarks["户数"].Text = cbfCount.ToString();
                exportWord.Range.Bookmarks["宗地"].Text = fieldCount.ToString();
                exportWord.Range.Bookmarks["面积"].Text = area.ToString("f1");
                exportWord.Range.Bookmarks["发包方地址2"].Text = fbfdz;
                exportWord.Save(targetFilePath);
                flag = true;
            }
            catch (Exception)
            {
                flag = false;
            }
            return flag;

        }

        /// <summary>
        /// 从发包方地址名称中获取村的名称
        /// </summary>
        /// <param name="fbfdz"></param>
        /// <returns></returns>
        private string SplitNameFromFbfdz(string fbfdz)
        {
            var zhenIndex = fbfdz.IndexOf("镇") == -1 ? fbfdz.IndexOf("乡") : fbfdz.IndexOf("镇");
            return fbfdz.Substring(zhenIndex + 1);
        }
        /// <summary>
        /// 获取发包方地址
        /// </summary>
        /// <returns>发包方地址</returns>
        private string GetFbfdz()
        {
            return Fbf().Fbfdz;
        }


        /// <summary>
        /// 获取宗地数目
        /// </summary>
        /// <returns></returns>
        private int GetFieldCount()
        {
            var cbfs = Cbfs(false);
            int count = 0;
            foreach (var cbfModel in cbfs)
            {
                 var fields = Fields(cbfModel.Cbfbm);
                count += fields.Count;
            }

            return count;
        }

        /// <summary>
        /// 获取总面积
        /// </summary>
        /// <returns></returns>
        private double? GetArea()
        {

            
            double area = 0.0;
            
            var cbfs = Cbfs(false);
            foreach (var cbf in cbfs)
            {
                var fields = Fields(cbf.Cbfbm);
                foreach (var field in fields)
                {
                    area += field.Scmj;
                }
            }
            return area;
            /*
            var query = from cbf in Cbfs(false)
                from field in Fields(cbf.Cbfbm)
                select field.Scmj;
            return query.Sum();
             *  */
        }
    }
}
