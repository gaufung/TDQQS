using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using ESRI.ArcGIS.Geodatabase;
using TdqqClient.Services.AE;
using TdqqClient.Services.Check;
using TdqqClient.Services.Database;
using TdqqClient.ViewModels;
using TdqqClient.Views;

namespace TdqqClient.Models.Edit
{
    /// <summary>
    /// 设置地块编码
    /// </summary>
    class EditDkbm:EditModel
    {
        public EditDkbm(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }

        public override void Edit(object parameter)
        {
            IAeFactory pAeFactory = new PersonalGeoDatabase(PersonDatabase);
            IFeatureClass pFeatureClass = pAeFactory.OpenFeatureClasss(SelectFeature);
            if (!pFeatureClass.FieldExistCheck("DKBM"))
            {
                System.Windows.Forms.MessageBox.Show(null,
                   "缺少地块编码字段", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            var dks = pFeatureClass.GetDks();
            var nsSortDks = dks.OrderBy(feature => feature.Ycor);
            DkbmViewModel dkbmVm = new DkbmViewModel(nsSortDks.Last().Ycor - nsSortDks.First().Ycor);
            DkbmView dkbmView = new DkbmView(dkbmVm);
            dkbmView.ShowDialog();
            if (dkbmVm.IsConfirm)
            {
                if (Dkbm(nsSortDks.Last().Ycor, nsSortDks.First().Ycor, dkbmVm.RowGap, dkbmVm.Fbfbm))
                {
                    System.Windows.Forms.MessageBox.Show(null,
                        "地块编码设置成功", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show(null,
                  "地块编码设置失败", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            } 
            pAeFactory.ReleaseFeautureClass(pFeatureClass);
        }

        private bool Dkbm(double yMax, double yMin, double rowGap, string fbfbm)
        {
            Wait wait = new Wait();
            wait.SetWaitCaption("设置地块编码");
            Hashtable para = new Hashtable()
            {
                {"wait",wait},
                {"yMax",yMax},
                {"yMin",yMin},
                {"rowGap",rowGap},
                {"fbfbm",fbfbm},
                {"ret",false}
            };
            Thread t = new Thread(new ParameterizedThreadStart(Dkbm));
            t.Start(para);
            wait.ShowDialog();
            t.Abort();
            return (bool)para["ret"];
        }

        private void Dkbm(object p)
        {
            Hashtable para = p as Hashtable;
            Wait wait = para["wait"] as Wait;
            var currentTop = (double)para["yMax"] + 10;
            var currentButtom = currentTop - (double)para["rowGap"];
            var currentIndex = 0;
            IAeFactory pAeFactory = new PersonalGeoDatabase(PersonDatabase);
            IFeatureClass pFeatureClass = pAeFactory.OpenFeatureClasss(SelectFeature);
            try
            {
                var count = pFeatureClass.Count();
                var dks = pFeatureClass.GetDks();
                while (currentTop > (double)para["yMin"] - 10)
                {
                    var currenDks = GetYDks(dks, currentTop, currentButtom);
                    Dkbm(currenDks, para["fbfbm"].ToString(), ref currentIndex, wait, count);
                    currentTop -= (double)para["rowGap"];
                    currentButtom -= (double)para["rowGap"];
                }
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

        private void Dkbm(IEnumerable<SortEnity<object>> currenDks, string fbfbm,
            ref int currentIndex, Wait wait, int count)
        {
            IDatabaseService pDatabaseService = new MsAccessDatabase(PersonDatabase);
            var xSortdk = currenDks.OrderBy(dk => dk.Xcor);
            var sqlString = string.Empty;
            foreach (var sortEnity in xSortdk)
            {
                currentIndex++;
                wait.SetProgress(((double)currentIndex / (double)count));
                sqlString = string.Format("update {0} set DKBM ='{1}' where OBJECTID = {2} ", SelectFeature,
                    fbfbm + currentIndex.ToString("00000"), sortEnity.Id);
                pDatabaseService.Execute(sqlString);
            }
        }

        private IEnumerable<SortEnity<object>> GetYDks(IEnumerable<SortEnity<object>> dks, double currentTop, double currentButtom)
        {
            foreach (var sortEnity in dks)
            {
                if (sortEnity.Ycor < currentTop && sortEnity.Ycor >= currentButtom)
                {
                    yield return sortEnity;
                }
            }
        }   
    }
}
