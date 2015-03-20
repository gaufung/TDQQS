using System;
using System.Collections;
using System.Data.OleDb;
using System.IO;
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
            var cbfInfoVm=new CbfInfoViewModel();
            var cbfInfoV=new CbfInfoView(cbfInfoVm);
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
            var wait=new Wait();
            wait.SetWaitCaption("正在导入承包方基础信息表");
            var para=new Hashtable()
            {
                {"wait",wait},{"openFilePath",openFilePath},{"cbfInfoVm",cbfInfoVm},{"ret",false}
            };
            var t=new Thread(new ParameterizedThreadStart(ImportF));
            t.Start(para);
            wait.ShowDialog();
            t.Abort();
            return (bool) para["ret"];
        }

        private void ImportF(object p)
        {
            var para = p as Hashtable;
            var wait = para["wait"] as Wait;
            var openFilePath = para["openFilePath"].ToString();
            var cbfInfoVm = para["cbfInfoVm"] as CbfInfoViewModel;
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
                using (var fileStream = new FileStream(openFilePath, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbookSource = new HSSFWorkbook(fileStream);
                    //先填写第一个sheet内容
                    ISheet sheetSource = (HSSFSheet)workbookSource.GetSheetAt(0);
                    int sheetRowCount = sheetSource.LastRowNum;
                    int startRowIndex = 1;
                    IRow rowSource = (HSSFRow)sheetSource.GetRow(startRowIndex);
                    ICell cell = null;
                    int currentIndex = 0;
                    while (rowSource != null)
                    {
                        wait.SetProgress(((double)currentIndex++ / (double)sheetRowCount));
                        string errorInfo = string.Empty;
                        if (!rowSource.IsDataRowValid(ref errorInfo))
                        {
                            MessageBox.Show(null, string.Format("第{0}行{1}", currentIndex + 1, errorInfo), 
                                "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return false;
                        }
                        CbfjtcyModel cbfjtcyModel = GetCbfjtcy(rowSource);
                        //往数据库中插入记录
                        var sqlString =
                            string.Format("insert into {0} (CBFBM,CYXB,CYXM,CYZJHM,CYZJLX,CYBZ,YHZGX,CYSZC,SFGYR,LXDH,YZBM) " +
                                "VALUES ('{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}')",
                                "CBF_JTCY", cbfjtcyModel.Cbfbm, cbfjtcyModel.Cyxb, cbfjtcyModel.Cyxm, cbfjtcyModel.Cyzjhm, 
                                cbfjtcyModel.Cyzjlx, cbfjtcyModel.Cybz, cbfjtcyModel.Yhzgx, cbfjtcyModel.Cyszc, 
                                cbfjtcyModel.Sfgyr, cbfjtcyModel.Lxdh, cbfjtcyModel.Yzbm);
                        if (!InsertRow(sqlString))
                        {
                            MessageBox.Show(null, string.Format("第{0} 行数据格式有误！", currentIndex + 1), "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return false;
                        }
                        startRowIndex++;
                        rowSource = sheetSource.GetRow(startRowIndex);
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

        /// <summary>
        /// 根据Excel中一行数据，获取一个CBFJTCY的一个对象
        /// </summary>
        /// <param name="rowSource"></param>
        /// <returns></returns>
        private CbfjtcyModel GetCbfjtcy(IRow rowSource)
        {
            const string defaultYzbm = "272600";
            return  new CbfjtcyModel()
            {
                Cbfbm = rowSource.GetCell(0).ToString().Trim(),
                Cyxb = rowSource.GetCell(1).ToString().Trim(),
                Cyxm =  rowSource.GetCell(2).ToString().Trim(),
                Cyzjlx = rowSource.GetCell(3).ToString(),
                Cyzjhm = rowSource.GetCell(4) == null ? string.Empty : rowSource.GetCell(4).ToString().Trim(),
                Cybz = rowSource.GetCell(5) == null ? string.Empty : rowSource.GetCell(5).ToString().Trim(),
                Yhzgx = rowSource.GetCell(6).ToString().Trim(),
                Cyszc = rowSource.GetCell(7) == null ? string.Empty : rowSource.GetCell(7).ToString().Trim(),
                Yzbm = rowSource.GetCell(8) == null ? defaultYzbm : rowSource.GetCell(8).ToString().Trim(),
                Sfgyr = rowSource.GetCell(9) == null ? string.Empty : rowSource.GetCell(9).ToString().Trim(),
                Lxdh = rowSource.GetCell(10) == null ? string.Empty : rowSource.GetCell(10).ToString().Trim()
            };
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
                IDatabaseService pDatabaseService = new MsAccessDatabase(BasicDatabase);
                var dt = pDatabaseService.Query(sqlString);
                if (dt == null) return false;
                int rowCount = dt.Rows.Count;
                int currentIndex = 0;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    wait.SetProgress( ((double)currentIndex++ / (double)rowCount));
                    var cbfbm = dt.Rows[i][0].ToString();
                    sqlString = string.Format("Select CBFBM from {0} where trim(CBFBM)='{1}'", "CBF_JTCY", cbfbm);
                    var dt1 = pDatabaseService.Query(sqlString);
                    var cbfModel = GetCbfModel(dt.Rows[i], cbfInfoVm);
                    cbfModel.Cbfcysl = dt1.Rows.Count;
                    const string queryString = "Insert Into [CBF] ([CBFBM],[CBFLX],[CBFMC],[CYXB],[CBFZJLX],[CBFZJHM],[CBFDZ],[YZBM],[LXDH],[CBFCYSL]," +
                                         "[CBFDCY],[CBFDCRQ],[CBFDCJS],[GSJS],[GSJSR],[GSSHR],[GSSHRQ]) " +
                                 "Values(@cbfbm,@cbflx,@cbfmc,@cyxb,@cbfzjlx,@cbfzjhm,@cbfdz,@yzbm,@lxdh,@cbfcysl,@cbfdcy,@cbfdcrq,@cbfdcjs,@gsjs,@gsjsr," +
                                         "@gsshr,@gsfhrq)";
                    var cn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;" + "data source=" + BasicDatabase);
                    cn.Open();
                    var cmd = new OleDbCommand(queryString, cn);                  
                    AddParameter(cmd,cbfModel);
                    cmd.ExecuteNonQuery();
                    cn.Close();
                    cn.Dispose();
                    dt1.Clear();
                }
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        private CbfModel GetCbfModel(System.Data.DataRow row, CbfInfoViewModel cbfInfoVm)
        {
            return new CbfModel()
             {
                 Cbfbm = row[0].ToString().Trim(),
                 Cbflx = "1",
                 Cbfmc = row[1].ToString().Trim(),
                 Cbfzjlx = row[2].ToString().Trim(),
                 Cbfzjhm = row[3].ToString().Trim(),
                 Cbfdz = row[4].ToString().Trim(),
                 Yzbm = row[5].ToString().Trim(),
                 Lxdh = row[6].ToString().Trim(),
                 Cyxb = row[7].ToString().Trim(),
                 Cbfdcy = cbfInfoVm.Cbfdcy,
                 Cbfdcjs = cbfInfoVm.Dcjs,
                 Gsjs = cbfInfoVm.Gsjs,
                 Gsjsr = cbfInfoVm.Gsjsr,
                 Gsshr = cbfInfoVm.Gsshr,
                 Cbfdcrq = cbfInfoVm.Dcrq,
                 Gsshrq = cbfInfoVm.Shrq
             };
        }

        private void AddParameter(OleDbCommand cmd, CbfModel cbfmodel)
        {
            cmd.Parameters.AddWithValue("@cbfbm", cbfmodel.Cbfbm);
            cmd.Parameters.AddWithValue("@cbflx", cbfmodel.Cbflx);
            cmd.Parameters.AddWithValue("@cbfmc", cbfmodel.Cbfmc);
            cmd.Parameters.AddWithValue("@cyxb", cbfmodel.Cyxb);
            cmd.Parameters.AddWithValue("@cbfzjlx", cbfmodel.Cbfzjlx);
            cmd.Parameters.AddWithValue("@cbfzjhm", cbfmodel.Cbfzjhm);
            cmd.Parameters.AddWithValue("@cbfdz", cbfmodel.Cbfdz);
            cmd.Parameters.AddWithValue("@yzbm", cbfmodel.Yzbm);
            cmd.Parameters.AddWithValue("@lxdh", cbfmodel.Lxdh);
            cmd.Parameters.AddWithValue("@cbfcysl", cbfmodel.Cbfcysl);
            cmd.Parameters.AddWithValue("@cbfdcy", cbfmodel.Cbfdcy);
            //添加承包方调查日期和公示审核日期
            var parameter1 = new OleDbParameter();
            parameter1.OleDbType = OleDbType.DBDate;
            parameter1.Value = cbfmodel.Cbfdcrq;
            cmd.Parameters.Add(parameter1);
            cmd.Parameters.AddWithValue("@cbfdcjs", cbfmodel.Cbfdcjs);
            cmd.Parameters.AddWithValue("@gsjs", cbfmodel.Gsjs);
            cmd.Parameters.AddWithValue("@gsjsr", cbfmodel.Gsjsr);
            cmd.Parameters.AddWithValue("@gsshr", cbfmodel.Gsshr);
            var parameter2 = new OleDbParameter();
            parameter2.OleDbType = OleDbType.DBDate;
            parameter2.Value = cbfmodel.Gsshrq;
            cmd.Parameters.Add(parameter2);
        }
    }
}
