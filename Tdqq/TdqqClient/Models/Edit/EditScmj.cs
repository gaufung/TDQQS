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
    class EditScmj:EditModel
    {
        public EditScmj(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }

        public override void Edit(object parameter)
        {
            if (!CheckEditFieldsExist())
            {
                MessageBox.Show(null, "字段尚未完全添加", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (Scmj())
            {
                MessageBox.Show(null, "实测面积设置成功", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            else
            {
                MessageBox.Show(null, "实测面积设置失败", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }   
        }

        private bool Scmj()
        {
            Wait wait = new Wait();
            wait.SetWaitCaption("设置实测面积");
            Hashtable para = new Hashtable()
            {
                {"wait",wait},
                {"ret",false}
            };
            Thread t = new Thread(new ParameterizedThreadStart(Scmj));
            t.Start(para);
            wait.ShowDialog();
            t.Abort();
            return (bool)para["ret"];
        }

        private void Scmj(object p)
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
                pWorkspaceEdit.StartEditing(true);
                pWorkspaceEdit.StartEditOperation();
                IFeature pFeature = pFeatureCursor.NextFeature();
                int currentIndex = 0;
                int shapeAreaIndex = AeHelper.FindFieldIndexInLowerCapper(pFeatureClass, "SHAPE_Area");
                while (pFeature != null)
                {
                    wait.SetProgress(((double)currentIndex++ / (double)total));
                    var scmj = Convert.ToDouble(pFeature.get_Value(shapeAreaIndex).ToString()) / 666.6;
                    pFeature.set_Value(pFeatureClass.FindField("SCMJ"), scmj);
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
                pAeFactory.ReleaseFeautureClass(pFeatureClass);
                wait.CloseWait();
            }
        }
    }
}
