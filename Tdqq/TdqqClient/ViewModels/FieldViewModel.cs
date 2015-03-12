using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TdqqClient.Commands;
using TdqqClient.Services.Database;

namespace TdqqClient.ViewModels
{
    public class FieldViewModel:NotificationObject
    {
        #region 属性字段
        private string _cbfmc;

        public string Cbfmc
        {
            get { return _cbfmc; }
            set
            {
                _cbfmc = value;
                this.RaisePropertyChanged("Cbfmc");
            }
        }
        private string _cbfbm;

        public string Cbfbm
        {
            get { return _cbfbm; }
            set
            {
                _cbfbm = value;
                this.RaisePropertyChanged("Cbfbm");
            }
        }

        private string _dkmc;

        public string Dkmc
        {
            get { return _dkmc; }
            set
            {
                _dkmc = value;
                this.RaisePropertyChanged("Dkmc");
            }
        }
        private string _dkbm;

        public string Dkbm
        {
            get { return _dkbm; }
            set { _dkbm = value; }
        }
        private double _htmj;

        public double Htmj
        {
            get { return _htmj; }
            set
            {
                _htmj = value;
                this.RaisePropertyChanged("Htmj");
            }
        }
        private double _yhtmj;

        public double Yhtmj
        {
            get { return _yhtmj; }
            set
            {
                _yhtmj = value;
                this.RaisePropertyChanged("Yhtmj");
            }
        }
        private double _scmj;

        public double Scmj
        {
            get { return _scmj; }
            set
            {
                _scmj = value;
                this.RaisePropertyChanged("Scmj");
            }
        }
        private string _dkdz;

        public string Dkdz
        {
            get { return _dkdz; }
            set
            {
                _dkdz = value;
                this.RaisePropertyChanged("Dkdz");
            }
        }

        private string _dknz;

        public string Dknz
        {
            get { return _dknz; }
            set
            {
                _dknz = value;
                this.RaisePropertyChanged("Dknz");
            }
        }
        private string _dkxz;

        public string Dkxz
        {
            get { return _dkxz; }
            set
            {
                _dkxz = value;
                this.RaisePropertyChanged("Dkxz");
            }
        }
        private string _dkbz;

        public string Dkbz
        {
            get { return _dkbz; }
            set
            {
                _dkbz = value;
                this.RaisePropertyChanged("Dkbz");
            }
        }
        
        #endregion

        #region 移动和关闭窗口

        public DelegateCommand WindowMoveCommand { get; set; }
        public DelegateCommand CloseCommand { get; set; }

        private void MoveWindow(object parameter)
        {
            this.OnMovingRequest();
        }

        private void CloseWindow(object parameter)
        {
            this.OnClosingRequest();
        }
        #endregion

        #region 保存信息

        public DelegateCommand ConfirmCommand { get; set; }

        private void Save(object parameter)
        {
            if (string.IsNullOrEmpty(Dkbm)) return;
            IDatabaseService pDatabaseService=new MsAccessDatabase(_personDatabase);
            string sqlString = string.Format("Update {0} Set CBFMC='{1}', DKMC='{2}', DKDZ='{3}'," +
                                             "DKNZ='{4}',DKXZ='{5}',DKBZ='{6}',YHTMJ={7},HTMJ={8},SCMJ={9} Where DKBM='{10}'",
                _selectFeauture,
                Cbfmc, Dkmc, Dkdz, Dknz, Dkxz, Dkbz, Yhtmj, Htmj, Scmj, Dkbm);
            var ret = pDatabaseService.Execute(sqlString);
            if (ret != -1)
            {
                MessageBox.Show(null, "保存成功", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(null, "保存失败", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        private readonly string _personDatabase;
        private readonly string _selectFeauture;

        public FieldViewModel(string personDatabase,string selectFeature)
        {
            _personDatabase = personDatabase;
            _selectFeauture = selectFeature;
            InitCommand();
        }

        private void InitCommand()
        {
            WindowMoveCommand=new DelegateCommand();
            CloseCommand=new DelegateCommand();
            ConfirmCommand=new DelegateCommand();
            WindowMoveCommand.ExecuteAction=new Action<object>(MoveWindow);
            CloseCommand.ExecuteAction=new Action<object>(CloseWindow);
            ConfirmCommand.ExecuteAction=new Action<object>(Save);

        }
    }
}
