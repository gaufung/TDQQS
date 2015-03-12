using System.Windows.Forms;
using TdqqClient.Services.Database;
using TdqqClient.ViewModels;
using TdqqClient.Views;

namespace TdqqClient.Models.Edit
{
    /// <summary>
    /// 设置默认值
    /// </summary>
    class EditSetDefault:EditModel
    {
        public EditSetDefault(string personDatabase, string selectFeauture, string basicDatabase)
            : base(personDatabase, selectFeauture, basicDatabase)
        {   }

        public override void Edit(object parameter)
        {
            if (!CheckEditFieldsExist())
            {
                MessageBox.Show(null, "字段尚未添加", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            var setDefaultVm = new SetDefaultViewModel();
            var setDefaultView = new SetDefaultView(setDefaultVm);
            setDefaultView.ShowDialog();
            if (setDefaultVm.IsConfirm)
            {
                if (SetDefaultField(setDefaultVm))
                {
                    MessageBox.Show(null, "字段设值成功", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(null, "字段设值失败", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }   
        }
        /// <summary>
        /// 设置默认值  Where OBJECTID <> {9}
        /// </summary>
        /// <param name="setDefautlVm"></param>
        /// <returns></returns>
        private bool SetDefaultField(SetDefaultViewModel setDefautlVm)
        {
            IDatabaseService pDatabaseService=new MsAccessDatabase(PersonDatabase);
            string sqlString = string.Format("Update {0} Set ZJRXM= '{1}',TDLYLX='{2}', CBJYQQDFS='{3}', SYQXZ='{4}'"
                + ",TDYT ='{5}', DLDJ ='{6}', SFJBNT='{7}', DKLB ='{8}' ", SelectFeature, setDefautlVm.Zjrxm, setDefautlVm.Tdlylx.Code,
                setDefautlVm.Cbjyqqdfs.Code,setDefautlVm.Syqxz.Code,setDefautlVm.Tdyt.Code,setDefautlVm.Dldj.Code,setDefautlVm.Sfjbnt.Code,setDefautlVm.Dklb.Code);
            var ret = pDatabaseService.Execute(sqlString);
            return ret == -1 ? false : true;            
        }      
    }
}
