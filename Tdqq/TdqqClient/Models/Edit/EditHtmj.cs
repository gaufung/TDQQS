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
using TdqqClient.Views;

namespace TdqqClient.Models.Edit
{
    /// <summary>
    /// 设置合同面积
    /// </summary>
    class EditHtmj:EditModel
    {
        public EditHtmj(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }

        public override void Edit(object parameter)
        {
            //base.Edit(parameter);
            if (!CheckEditFieldsExist())
            {
                MessageBox.Show(null, "字段尚未完全添加", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (Htmj())
            {
                MessageBox.Show(null, "合同面积设置成功", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(null, "合同面积设置失败", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool Htmj()
        {
            Wait wait = new Wait();
            wait.SetWaitCaption("设置合同面积");
            Hashtable para = new Hashtable()
            {
                {"wait",wait},
                {"ret",false}
            };
            Thread t = new Thread(new ParameterizedThreadStart(Htmj));
            t.Start(para);
            wait.ShowDialog();
            t.Abort();
            return (bool)para["ret"];
        }

        private void Htmj(object p)
        {
            Hashtable para = p as Hashtable;
            Wait wait = para["wait"] as Wait;
            IAeFactory pAeFactory = new PersonalGeoDatabase(PersonDatabase);
            IFeatureClass pFeatureClass = pAeFactory.OpenFeatureClasss(SelectFeature);
            try
            {
                var pDataset = pFeatureClass as IDataset;
                var pWorkspaceEdit = pDataset.Workspace as IWorkspaceEdit;
                IFeatureCursor pFeatureCursor = pFeatureClass.Search(null, false);
                int total = pFeatureClass.Count();
                int shapeAreaIndex = pFeatureClass.FindFieldIndexInLowerCapper("SHAPE_Area");
                pWorkspaceEdit.StartEditing(true);
                pWorkspaceEdit.StartEditOperation();
                IFeature pFeature;
                int currentIndex = 0;
                while ((pFeature = pFeatureCursor.NextFeature()) != null)
                {
                    wait.SetProgress(((double)currentIndex++ / (double)total));
                    if (pFeatureClass.Fields.FindField("YHTMJ") == -1 ||
                        string.IsNullOrEmpty(
                            pFeature.Value[pFeatureClass.Fields.FindField("YHTMJ")].ToString().Trim()))
                    {
                        var scmj = Convert.ToDouble(pFeature.Value[shapeAreaIndex].ToString()) /
                                   666.6;
                        pFeature.Value[pFeatureClass.FindField("HTMJ")]=scmj;
                    }
                    else
                    {
                        pFeature.Value[pFeatureClass.FindField("HTMJ")]=
                            pFeature.Value[pFeatureClass.FindField("YHTMJ")];
                    }
                    pFeature.Store();
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
                wait.CloseWait();
                pAeFactory.ReleaseFeautureClass(pFeatureClass);
            }
        }   
    }
}
