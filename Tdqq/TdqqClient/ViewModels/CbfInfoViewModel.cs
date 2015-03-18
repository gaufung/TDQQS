using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TdqqClient.Commands;

namespace TdqqClient.ViewModels
{
    public class CbfInfoViewModel:NotificationObject
    {
        public bool IsConfirm { get; set; }

        #region 关闭窗口

        public DelegateCommand CloseCommand { get; set; }

        private void CloseWindow(object parameter)
        {
            this.OnClosingRequest();
            IsConfirm = false;
        }
        #endregion

        #region 绑定属性

        private string _cbfdcy;

        public string Cbfdcy
        {
            get { return _cbfdcy; }
            set
            {
                _cbfdcy = value;
                this.RaisePropertyChanged("Cbfdcy");
            }
        }
        private DateTime _dcrq;

        public DateTime Dcrq
        {
            get { return _dcrq; }
            set
            {
                _dcrq = value;
                this.RaisePropertyChanged("Dcrq");
            }
        }
        private string _dcjs;

        public string Dcjs
        {
            get { return _dcjs; }
            set { _dcjs = value;this.RaisePropertyChanged("Dcjs"); }
        }
        private string _gsjs;

        public string Gsjs
        {
            get { return _gsjs; }
            set { _gsjs = value; this.RaisePropertyChanged("Gsjs");}
        }
        private DateTime _shrq;

        public DateTime Shrq
        {
            get { return _shrq; }
            set { _shrq = value; this.RaisePropertyChanged("Shrq");}
        }

        private string _gsjsr;

        public string Gsjsr
        {
            get { return _gsjsr; }
            set { _gsjsr = value; this.RaisePropertyChanged("");}
        }
        private string _gsshr;

        public string Gsshr
        {
            get { return _gsshr; }
            set { _gsshr = value; this.RaisePropertyChanged("Gsshr");}
        }
        
        #endregion

        #region 确定命令

        public DelegateCommand ConfirmCommand { get; set; }

        private void Confirm(object parameter)
        {
            this.IsConfirm = true;
           this.OnClosingRequest();
        }
        #endregion

        public CbfInfoViewModel()
        {
            this.CloseCommand=new DelegateCommand();
            this.ConfirmCommand=new DelegateCommand();
            this.CloseCommand.ExecuteAction=new Action<object>(CloseWindow);
            this.ConfirmCommand.ExecuteAction=new Action<object>(Confirm);
            this.Shrq=new DateTime(2014,10,18);
            this.Dcrq=new DateTime(2014,7,10);
            Cbfdcy = Dcjs = Gsjs = Gsjsr = Gsshr = string.Empty;
        }
    }
}
