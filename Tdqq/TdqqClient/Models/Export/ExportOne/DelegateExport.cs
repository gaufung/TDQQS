using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using Aspose.Words;
using TdqqClient.Services.Common;
using TdqqClient.Services.Database;
using TdqqClient.Views;

namespace TdqqClient.Models.Export.ExportOne
{
    class DelegateExport:ExportOne
    {
        public DelegateExport(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }

        public override void Export(object parameter)
        {
            var dialogHelper=new DialogHelper("doc");
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
            Wait wait=new Wait();
            wait.SetWaitCaption("导出委托书");
            Hashtable para=new Hashtable()
            {
                {"wait",wait},{"saveFilePath",saveFilePath},{"ret",false}
            };
            Thread t=new Thread(new ParameterizedThreadStart(ExportF));
            t.Start(para);
            wait.ShowDialog();
            t.Abort();
            return (bool) para["ret"];
        }

        private void ExportF(object p)
        {
            var para = p as Hashtable;
            var wait = para["wait"] as Wait;
            var saveDocPath = para["saveFilePath"].ToString().Trim();
            var templatePath = AppDomain.CurrentDomain.BaseDirectory + @"\template\村民委托书.doc";
            try
            {
                var cbfInfo = GetCbfInfo();
                string zhen, cun;
                zhen = cun = string.Empty;
                Fbfxx(ref zhen, ref cun);
                var count = cbfInfo.Count(a => a.Cbfmc != string.Empty);
                var sortCbfInfo = cbfInfo.OrderBy(cbf => cbf.Cbfmc);

                Document doc = new Document(templatePath);
                DocumentBuilder documentBuilder = new DocumentBuilder(doc);
                doc.Range.Bookmarks["zhen"].Text = zhen;
                doc.Range.Bookmarks["cun"].Text = cun;
                List<double> widthList = new List<double>();
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
                    wait.SetProgress(((double) index/(double) count));
                    for (int i = 0; i < columnCount; i++)
                    {
                        documentBuilder.InsertCell(); // 添加一个单元格                    
                        documentBuilder.CellFormat.Borders.LineStyle = Aspose.Words.LineStyle.Single;
                        documentBuilder.CellFormat.Borders.Color = System.Drawing.Color.Black;
                        documentBuilder.CellFormat.Width = widthList[i];
                        documentBuilder.CellFormat.VerticalMerge = Aspose.Words.Tables.CellMerge.None;
                        documentBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                        if (i == 0)
                        {
                            documentBuilder.Write(index.ToString());
                        }
                        if (i == 1)
                        {
                            documentBuilder.Write(info.Cbfmc);
                        }
                        if (i == 2)
                        {
                            documentBuilder.Write(info.Cbfzjhm);
                        }
                    }
                    index++;
                    documentBuilder.EndRow();
                }
                doc.Save(saveDocPath, SaveFormat.Doc);
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
        private IEnumerable<FarmerModel> GetCbfInfo()
        {
            var sqlString = string.Format("Select distinct CBFBM,CBFMC From {0} where CBFBM NOT LIKE  '{1}' order by CBFBM ",
              SelectFeature, "99999999999999%");
            var accessFactory = new MsAccessDatabase(PersonDatabase);
            var dt = accessFactory.Query(sqlString);
            accessFactory = new MsAccessDatabase(BasicDatabase);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                sqlString = string.Format("select CBFZJHM from {0} where trim(CBFBM)='{1}'", "CBF",
                    dt.Rows[i][0].ToString().Trim());
                var dtZjhm = accessFactory.Query(sqlString);
                yield return new FarmerModel()
                {
                    Cbfmc = dt.Rows[i][1].ToString(),
                    Cbfzjhm = dtZjhm.Rows[0][0].ToString()
                };
            }
        }
        private void Fbfxx(ref string zhen, ref string cun)
        {
            var sqlString = string.Format("Select FBFDZ from FBF");
            var accessFactory = new MsAccessDatabase(BasicDatabase);
            var dt = accessFactory.Query(sqlString);
            if (dt == null || dt.Rows.Count != 1) return;
            var fbfdz = dt.Rows[0][0].ToString();
            var xianIndex = fbfdz.IndexOf("县");
            var zhenIndex = fbfdz.IndexOf("镇") == -1 ? fbfdz.IndexOf("乡") : fbfdz.IndexOf("镇");
            var cunIndex = fbfdz.IndexOf("村");
            zhen = fbfdz.Substring(xianIndex + 1, zhenIndex - xianIndex);
            cun = fbfdz.Substring(zhenIndex + 1, cunIndex - zhenIndex - 1);
        }
    }
}
