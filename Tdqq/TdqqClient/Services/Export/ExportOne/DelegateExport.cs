using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using Aspose.Words;
using TdqqClient.Services.Common;
using TdqqClient.Views;

namespace TdqqClient.Services.Export.ExportOne
{
    /// <summary>
    /// 导出村民委托书
    /// </summary>
    class DelegateExport:ExportBase,IExport
    {
        public DelegateExport(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }
        public void Export()
        {
            var dialogHelper = new DialogHelper("doc");
            var saveFilePath = dialogHelper.SaveFile("导出委托书");
            if (string.IsNullOrEmpty(saveFilePath)) return;
            if (Export(saveFilePath))
            {
                MessageBox.Show(null, "委托书导出成功",
                    "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(null, "委托书导出失败", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool Export(string saveFilePath)
        {
            var wait = new Wait();
            wait.SetWaitCaption("导出委托书");
            var para = new Hashtable()
            {
                {"wait",wait},{"saveFilePath",saveFilePath},{"ret",false}
            };
            var t = new Thread(new ParameterizedThreadStart(ExportF));
            t.Start(para);
            wait.ShowDialog();
            t.Abort();
            return (bool)para["ret"];
        }

        /// <summary>
        /// 另外开启一个线程用来导出委托书
        /// </summary>
        /// <param name="p"></param>
        private void ExportF(object p)
        {
            var para = p as Hashtable;
            var wait = para["wait"] as Wait;
            var saveDocPath = para["saveFilePath"].ToString().Trim();
            var templatePath = AppDomain.CurrentDomain.BaseDirectory + @"\template\村民委托书.doc";
            try
            {
                var cbfInfo = Cbfs(true);
                string zhen, cun;
                zhen = cun = string.Empty;
                Fbfxx(ref zhen, ref cun);
                var count = cbfInfo.Count(a => a.Cbfmc != string.Empty);
                var sortCbfInfo = cbfInfo.OrderBy(cbf => cbf.Cbfmc);
                var doc = new Document(templatePath);
                var documentBuilder = new DocumentBuilder(doc);
                doc.Range.Bookmarks["zhen"].Text = zhen;
                doc.Range.Bookmarks["cun"].Text = cun;
                var widthList = new List<double>();
                //获取列宽
                const int columnCount = 4;
                for (int i = 0; i < columnCount; i++)
                {
                    documentBuilder.MoveToCell(0, 0, i, 0); //移动单元格
                    double width = documentBuilder.CellFormat.Width; //获取单元格宽度
                    widthList.Add(width);
                }
                documentBuilder.MoveToBookmark("table");
                int index = 1;
                foreach (var info in sortCbfInfo)
                {
                    wait.SetProgress(((double)index / (double)count));
                    for (int i = 0; i < columnCount; i++)
                    {
                        documentBuilder.InsertCell(); // 添加一个单元格                    
                        documentBuilder.CellFormat.Borders.LineStyle = Aspose.Words.LineStyle.Single;
                        documentBuilder.CellFormat.Borders.Color = System.Drawing.Color.Black;
                        documentBuilder.CellFormat.Width = widthList[i];
                        documentBuilder.CellFormat.VerticalMerge = Aspose.Words.Tables.CellMerge.None;
                        documentBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                        if (i == 0) documentBuilder.Write(index.ToString());
                        if (i == 1) documentBuilder.Write(info.Cbfmc);
                        if (i == 2) documentBuilder.Write(info.Cbfzjhm);
                    }
                    index++;
                    documentBuilder.EndRow();
                }
                doc.Save(saveDocPath, SaveFormat.Doc);
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
        /// <summary>
        /// 获取发包方相关信息
        /// </summary>
        /// <param name="zhen">镇</param>
        /// <param name="cun">村</param>
        private void Fbfxx(ref string zhen, ref string cun)
        {
            var fbf = Fbf();
            if (fbf==null) return;
            var fbfdz = fbf.Fbfdz;
            var xianIndex = fbfdz.IndexOf("县");
            var zhenIndex = fbfdz.IndexOf("镇") == -1 ? fbfdz.IndexOf("乡") : fbfdz.IndexOf("镇");
            var cunIndex = fbfdz.IndexOf("村");
            zhen = fbfdz.Substring(xianIndex + 1, zhenIndex - xianIndex);
            cun = fbfdz.Substring(zhenIndex + 1, cunIndex - zhenIndex - 1);
        }

    }
}
