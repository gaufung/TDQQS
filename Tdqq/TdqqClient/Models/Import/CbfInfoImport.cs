using System;
using System.Collections;
using System.Data.OleDb;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using TdqqClient.Services.Common;
using TdqqClient.Services.Database;
using TdqqClient.ViewModels;
using TdqqClient.Views;
using TdqqClient.Services.Check;

namespace TdqqClient.Models.Import
{
    /// <summary>
    /// 承包方家庭家庭成员信息导入
    /// </summary>
    class CbfInfoImport:ImportBase
    {



        public CbfInfoImport(string basicDatabase):base(basicDatabase)
        {            
        }

        public override void Import()
        {
            var dialogHelper=new DialogHelper("xls");
            var openFilePath = dialogHelper.OpenFile("选择基础信息表");
            if (string.IsNullOrEmpty(openFilePath)) return;
            if (!ValidCheck.ExcelColumnSorted(openFilePath))
            {
                MessageBox.Show(null, "基础信息表列顺序不满足要求", 
                    "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            CbfInfoViewModel cbfInfoVm=new CbfInfoViewModel();
            CbfInfoView cbfInfoV=new CbfInfoView(cbfInfoVm);
            cbfInfoV.ShowDialog();
            if (cbfInfoVm.IsConfirm)
            {
                if (Import(openFilePath, cbfInfoVm))
                {
                    MessageBox.Show(null, "信息导入成功",
                        "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(null, "信息导入失败", 
                        "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private bool Import(string openFilePath, CbfInfoViewModel cbfInfoVm)
        {
            Wait wait=new Wait();
            wait.SetWaitCaption("正在导入承包方基础信息表");
            Hashtable para=new Hashtable()
            {
                {"wait",wait},{"openFilePath",openFilePath},{"cbfInfoVm",cbfInfoVm},{"ret",false}
            };
            Thread t=new Thread(new ParameterizedThreadStart(ImportF));
            t.Start(para);
            wait.ShowDialog();
            t.Abort();
            return (bool) para["ret"];
        }

        private void ImportF(object p)
        {
            Hashtable para = p as Hashtable;
            Wait wait = para["wait"] as Wait;
            var openFilePath = para["openFilePath"].ToString();
            CbfInfoViewModel cbfInfoVm = para["cbfInfoVm"] as CbfInfoViewModel;
            if (!ImportCbfjtcy(openFilePath,wait))
            {
                wait.CloseWait();
                para["ret"] = false;
                return;
            }
            wait.SetWaitCaption("提取承包方名称");
            if (!SetCbfmc(wait))
            {
                wait.CloseWait();
                para["ret"] = false;
                return;
            }
            wait.SetWaitCaption("创建承包方表");
            if (!CreateCbf(wait,cbfInfoVm))
            {
                wait.CloseWait();
                para["ret"] = false;
                return;
            }
            wait.CloseWait();
            para["ret"] = true;


        }

        private bool ImportCbfjtcy(string openFilePath, Wait wait)
        {
            try
            {
                if (!DeleteTable("CBF_JTCY")) return false;
                using (FileStream fileStream = new FileStream(openFilePath, FileMode.Open, FileAccess.Read))
                {
                    HSSFWorkbook workbookSource = new HSSFWorkbook(fileStream);
                    //先填写第一个sheet内容
                    HSSFSheet sheetSource = (HSSFSheet)workbookSource.GetSheetAt(0);
                    int sheetRowCount = sheetSource.LastRowNum;
                    int start_row_index = 1;
                    HSSFRow rowSource = (HSSFRow)sheetSource.GetRow(start_row_index);
                    ICell cell = null;
                    int currentIndex = 0;
                    StringBuilder stringBuilder = new StringBuilder();
                    stringBuilder.Append("导入家庭成员信息错误信息：\r\n");
                    while (rowSource != null)
                    {
                        wait.SetProgress(((double)currentIndex++ / (double)sheetRowCount));
                        string errorInfo = string.Empty;
                        //if (!ValidCheck.(rowSource, ref errorInfo))
                        if (!rowSource.IsDataRowValid(ref errorInfo))
                        {
                            MessageBox.Show(null, string.Format("第{0}行{1}", currentIndex + 1, errorInfo), 
                                "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return false;
                        }
                        var cbfbm = rowSource.GetCell(0).ToString().Trim();
                        var cyxb = rowSource.GetCell(1).ToString().Trim();
                        var cyxm = rowSource.GetCell(2).ToString().Trim();
                        var cyzjlx = rowSource.GetCell(3).ToString();
                        var cyzjhm = rowSource.GetCell(4) == null ? string.Empty : rowSource.GetCell(4).ToString().Trim();
                        var cybz = rowSource.GetCell(5) == null ? string.Empty : rowSource.GetCell(5).ToString().Trim();
                        var yhzgx = rowSource.GetCell(6).ToString().Trim();
                        var cyszc = rowSource.GetCell(7) == null ? string.Empty : rowSource.GetCell(7).ToString().Trim();
                        var yzbm = rowSource.GetCell(8) == null ? "272600" : rowSource.GetCell(8).ToString().Trim();
                        var sfgyr = rowSource.GetCell(9) == null ? string.Empty : rowSource.GetCell(9).ToString().Trim();
                        var lxdh = rowSource.GetCell(10) == null ? string.Empty : rowSource.GetCell(10).ToString().Trim();
                        //往数据库中插入记录
                        var sqlString =
                            string.Format("insert into {0} (CBFBM,CYXB,CYXM,CYZJHM,CYZJLX,CYBZ,YHZGX,CYSZC,SFGYR,LXDH,YZBM) " +
                                "VALUES ('{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}')",
                                "CBF_JTCY", cbfbm, cyxb, cyxm, cyzjhm, cyzjlx, cybz, yhzgx, cyszc, sfgyr, lxdh, yzbm);
                        if (!InsertRow(sqlString))
                        {
                            System.Windows.Forms.MessageBox.Show(string.Format("第{0} 行数据格式有误！", currentIndex + 1));
                            return false;
                        }
                        start_row_index++;
                        rowSource = (HSSFRow)sheetSource.GetRow(start_row_index);
                    }
                    fileStream.Close();
                }
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        private bool SetCbfmc(Wait wait)
        {
            try
            {
                var sqlString = string.Format("update CBF_JTCY set CBFMC = CYXM where trim(YHZGX)='02'");
                var accessFactory = new MsAccessDatabase(BasicDatabase);
                var ret = accessFactory.Execute(sqlString);
                if (ret == -1) return false;
                sqlString = string.Format("select CBFBM,CBFMC from CBF_JTCY where trim(YHZGX)='02'");
                var dt = accessFactory.Query(sqlString);
                if (dt == null) return false;
                int rowCount = dt.Rows.Count;
                int currentIndex = 0;
                for (int i = 0; i < rowCount; i++)
                {
                    wait.SetProgress(((double)currentIndex++ / (double)rowCount));
                    var cbfmc = dt.Rows[i][1].ToString().Trim();
                    sqlString = string.Format("update CBF_JTCY set CBFMC='{0}'where trim(CBFBM)='{1}'", cbfmc,
                        dt.Rows[i][0].ToString().Trim());
                    accessFactory.Execute(sqlString);
                }
                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                return false;
            }
        }

        private bool CreateCbf(Wait wait,CbfInfoViewModel cbfInfoVm)
        {
            try
            {
                DeleteTable("CBF");
                var sqlString = string.Format("Select CBFBM,CBFMC,CYZJLX,CYZJHM,CYSZC,YZBM,LXDH,CYXB from {0} " +
                                              "Where trim(YHZGX)='{1}'", "CBF_JTCY", "02");
                IDatabaseService accessFactory = new MsAccessDatabase(BasicDatabase);
                var dt = accessFactory.Query(sqlString);
                if (dt == null) return false;
                int rowCount = dt.Rows.Count;
                int currentIndex = 0;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    wait.SetProgress( ((double)currentIndex++ / (double)rowCount));
                    var cbfbm = dt.Rows[i][0].ToString();
                    sqlString = string.Format("Select CBFBM from {0} where trim(CBFBM)='{1}'", "CBF_JTCY", cbfbm);
                    var dt1 = accessFactory.Query(sqlString);
                    var cbfcysl = dt1.Rows.Count;
                    var cbflx = "1";
                    var cbfmc = dt.Rows[i][1].ToString().Trim();
                    var cbfzjlx = dt.Rows[i][2].ToString().Trim();
                    var cbfzjhm = dt.Rows[i][3].ToString().Trim();
                    var cbfdz = dt.Rows[i][4].ToString().Trim();
                    var yzbm = dt.Rows[i][5].ToString().Trim();
                    var lxdh = dt.Rows[i][6].ToString().Trim();
                    var cyxb = dt.Rows[i][7].ToString().Trim();
                    string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" + "data source=" + BasicDatabase;
                    string queryString = "Insert Into [CBF] ([CBFBM],[CBFLX],[CBFMC],[CYXB],[CBFZJLX],[CBFZJHM],[CBFDZ],[YZBM],[LXDH],[CBFCYSL]," +
                                         "[CBFDCY],[CBFDCRQ],[CBFDCJS],[GSJS],[GSJSR],[GSSHR],[GSSHRQ]) " +
                                 "Values(@cbfbm,@cbflx,@cbfmc,@cyxb,@cbfzjlx,@cbfzjhm,@cbfdz,@yzbm,@lxdh,@cbfcysl,@cbfdcy,@cbfdcrq,@cbfdcjs,@gsjs,@gsjsr," +
                                         "@gsshr,@gsfhrq)";
                    var cn = new OleDbConnection(connectionString);
                    cn.Open();
                    var cmd = new OleDbCommand(queryString, cn);
                    cmd.Parameters.AddWithValue("@cbfbm", cbfbm);
                    cmd.Parameters.AddWithValue("@cbflx", cbflx);
                    cmd.Parameters.AddWithValue("@cbfmc", cbfmc);
                    cmd.Parameters.AddWithValue("@cyxb", cyxb);
                    cmd.Parameters.AddWithValue("@cbfzjlx", cbfzjlx);
                    cmd.Parameters.AddWithValue("@cbfzjhm", cbfzjhm);
                    cmd.Parameters.AddWithValue("@cbfdz", cbfdz);
                    cmd.Parameters.AddWithValue("@yzbm", yzbm);
                    cmd.Parameters.AddWithValue("@lxdh", lxdh);
                    cmd.Parameters.AddWithValue("@cbfcysl", cbfcysl);
                    cmd.Parameters.AddWithValue("@cbfdcy", cbfInfoVm.Cbfdcy);
                    //添加承包方调查日期和公示审核日期
                    var parameter1 = new OleDbParameter();
                    parameter1.OleDbType = OleDbType.DBDate;
                    parameter1.Value = cbfInfoVm.Dcrq;
                    cmd.Parameters.Add(parameter1);
                    cmd.Parameters.AddWithValue("@cbfdcjs", cbfInfoVm.Dcjs);
                    cmd.Parameters.AddWithValue("@gsjs", cbfInfoVm.Gsjs);
                    cmd.Parameters.AddWithValue("@gsjsr", cbfInfoVm.Gsjsr);
                    cmd.Parameters.AddWithValue("@gsshr", cbfInfoVm.Gsshr);
                    var parameter2 = new OleDbParameter();
                    parameter2.OleDbType = OleDbType.DBDate;
                    parameter2.Value = cbfInfoVm.Shrq;
                    cmd.Parameters.Add(parameter2);
                    cmd.ExecuteNonQuery();
                    cn.Close();
                    cn.Dispose();
                    dt1.Clear();
                }
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
