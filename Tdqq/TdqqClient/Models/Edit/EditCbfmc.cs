using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TdqqClient.Services.Database;
using TdqqClient.ViewModels;
using TdqqClient.Views;

namespace TdqqClient.Models.Edit
{
    class EditCbfmc:EditModel
    {
        public EditCbfmc(string personDatabase, string selectFeauture, string basicDatabase)
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
            ReplaceViewModel replaceVm = new ReplaceViewModel();
            ReplaceView replaceV = new ReplaceView(replaceVm);
            replaceV.ShowDialog();
            if (replaceVm.IsConfirm)
            {
                var ret = Replace(replaceVm);
                if (ret != -1)
                {
                    MessageBox.Show(null, string.Format("共替换{0}处承包方名称", ret), "信息提示", MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(null, "替换失败", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }       
        }

        private int Replace(ReplaceViewModel replaceVm)
        {
            IDatabaseService pDatabaseService = new MsAccessDatabase(PersonDatabase);
            var sqlString = string.Format("Update {0} Set CBFMC = '{1}' Where CBFMC ='{2}'", SelectFeature,
                 replaceVm.NewName, replaceVm.OriginalName);
            return pDatabaseService.Execute(sqlString);

        }   
    }
}
