using System;
using System.Data.OleDb;
using System.Windows.Forms;
using TdqqClient.ViewModels;
using TdqqClient.Views;

namespace TdqqClient.Models.Import
{
    /// <summary>
    /// 导入发包方信息
    /// </summary>
    class FbfInfoImport:ImportBase
    {
        public FbfInfoImport(string basicDatabase) : base(basicDatabase)
        {
        }

        public override void Import()
        {
            
            var fbfInfoVm = new FbfInfoViewModel();
            var fbfInfoV=new FbfInfoView(fbfInfoVm);
            fbfInfoV.ShowDialog();
            if (fbfInfoVm.IsConfirm)
            {
                if (Import(fbfInfoVm))
                {
                    MessageBox.Show(null, 
                        "发包方信息导入成功", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(null, "发包方信息导入失败", 
                        "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private bool Import(FbfInfoViewModel fbfInfoVm)
        {
            try
            {
                DeleteTable("FBF");
                var connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" + "data source=" + BasicDatabase;
                const string queryString = "Insert Into [FBF] ([FBFBM],[FBFMC],[FBFFZRXM],[FZRZJLX],[FZRZJHM],[LXDH],[FBFDZ],[YZBM],[FBFDCY],[FBFDCJS],[FBFDCRQ]) " +
                            "Values(@fbfbm, @fbfmc, @fbffzrxm, @fzrzjlx,@fzrzjhm,@lxdh,@fbfdz,@yzbm,@fbfdcy,@fbfdcjs,@fbffdrq)";
                var cn = new OleDbConnection(connectionString);
                cn.Open();
                var cmd = new OleDbCommand(queryString, cn);
                AddParameters(cmd,fbfInfoVm);
                var ret = cmd.ExecuteNonQuery();
                cn.Close();
                cn.Dispose();
                return ret == -1 ? false : true;
            }
            catch (Exception e)
            { 
                return false;
            }           
        }

        private void AddParameters(OleDbCommand cmd,FbfInfoViewModel fbfInfoVm)
        {
            cmd.Parameters.AddWithValue("@fbfbm", fbfInfoVm.Fbfbm);
            cmd.Parameters.AddWithValue("@fbfmc", fbfInfoVm.Fbfmc);
            cmd.Parameters.AddWithValue("@fbffzrxm", fbfInfoVm.Fzrxm);
            cmd.Parameters.AddWithValue("@fzrzjlx", fbfInfoVm.Fzrzjlx.Code);
            cmd.Parameters.AddWithValue("@fzrzjhm", fbfInfoVm.Zjhm);
            cmd.Parameters.AddWithValue("@lxdh", fbfInfoVm.Lxdh);
            cmd.Parameters.AddWithValue("@fbfdz", fbfInfoVm.Fbfdz);
            cmd.Parameters.AddWithValue("@yzbm", fbfInfoVm.Yzbm);
            cmd.Parameters.AddWithValue("@fbfdcy", fbfInfoVm.Dcy);
            cmd.Parameters.AddWithValue("@fbfdcjs", fbfInfoVm.Dcjs);
            var parameter = new OleDbParameter();
            parameter.OleDbType = OleDbType.DBDate;
            parameter.Value = fbfInfoVm.Dcrq;
            cmd.Parameters.Add(parameter);
        }
    }
}
