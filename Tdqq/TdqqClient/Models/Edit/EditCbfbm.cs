using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using ESRI.ArcGIS.Geodatabase;
using TdqqClient.Services.AE;
using TdqqClient.Services.Check;
using TdqqClient.Services.Database;
using TdqqClient.Views;

namespace TdqqClient.Models.Edit
{
    class EditCbfbm:EditModel
    {
        public EditCbfbm(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }


        public override void Edit(object parameter)
        {
            //base.Edit(parameter);
            if (!CheckEditFieldsExist())
            {
                MessageBox.Show(null, "字段尚未添加成功", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (Update())
            {
                MessageBox.Show(null, "承包方编码更新成功", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(null, "承包方编码更新失败", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }   
        }

        private bool Update()
        {
            Wait wait = new Wait();
            wait.SetWaitCaption("更新承包方编码");
            Hashtable para = new Hashtable()
            {
                {"wait",wait},
                {"ret",false}
            };
            Thread t = new Thread(new ParameterizedThreadStart(Update));
            t.Start(para);
            wait.ShowDialog();
            t.Abort();
            return (bool)para["ret"];
        }

        private void Update(object p)
        {
            Hashtable para = p as Hashtable;
            Wait wait = para["wait"] as Wait;
            IAeFactory aeFactory = new PersonalGeoDatabase(PersonDatabase);
            IFeatureClass pFeatureClass = aeFactory.OpenFeatureClasss(SelectFeature);
            var cbfmcIndex = pFeatureClass.Fields.FindField("CBFMC");
            var cbfbmIndex = pFeatureClass.Fields.FindField("CBFBM");
            try
            {
                var pDataset = pFeatureClass as IDataset;
                var pWorkspaceEdit = pDataset.Workspace as IWorkspaceEdit;
                IFeatureCursor pFeatureCursor = pFeatureClass.Search(null, false);
                pWorkspaceEdit.StartEditing(true);
                pWorkspaceEdit.StartEditOperation();
                IFeature pFeature = pFeatureCursor.NextFeature();
                int errorIndex = 0;
                int currentIndex = 0;
                int total = pFeatureClass.Count();
                Dictionary<string, string> errorCbfmcDictionary = new Dictionary<string, string>();
                while (pFeature != null)
                {
                    wait.SetProgress(((double)currentIndex++ / (double)total));
                    var cbfmc = pFeature.get_Value(cbfmcIndex).ToString().Trim();
                    if (errorCbfmcDictionary.ContainsKey(cbfmc))
                    {
                        string errorCbfbm;
                        errorCbfmcDictionary.TryGetValue(cbfmc, out errorCbfbm);
                        pFeature.set_Value(cbfbmIndex, errorCbfbm);
                    }
                    else
                    {
                        var cbfbm = GetCbfbm(cbfmc, ref errorIndex);
                        if (cbfbm.StartsWith("9999"))
                        {
                            errorCbfmcDictionary.Add(cbfmc, cbfbm);
                        }
                        pFeature.set_Value(cbfbmIndex, cbfbm);
                    }
                    pFeature.Store();
                    pFeature = pFeatureCursor.NextFeature();
                }
                Marshal.ReleaseComObject(pFeatureCursor);
                pWorkspaceEdit.StopEditOperation();
                pWorkspaceEdit.StopEditing(true);
                para["ret"] = true;
            }
            catch (Exception)
            {
                para["ret"] = false;
            }
            finally
            {
                aeFactory.ReleaseFeautureClass(pFeatureClass);
                wait.CloseWait();
            }
        }
        private string GetCbfbm(string cbfmc, ref int errorIndex)
        {
            var sqlString = string.Format("SELECT CBFBM FROM {0} WHERE TRIM(CBFMC)='{1}' ", "CBF", cbfmc);
            IDatabaseService pDatabaseService = new MsAccessDatabase(BasicDatabase);
            var dt = pDatabaseService.Query(sqlString);
            if (dt == null || dt.Rows.Count != 1)
            {
                errorIndex++;
                return "99999999999999" + errorIndex.ToString("0000");
            }
            else
            {
                return dt.Rows[0][0].ToString().Trim();
            }
        }   
    }
}
