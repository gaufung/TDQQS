using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms.VisualStyles;
using TdqqClient.Commands;
using TdqqClient.Models;

namespace TdqqClient.ViewModels
{
    public class FbfInfoViewModel:NotificationObject
    {
        public bool IsConfirm { get; set; }
        #region 关闭窗口

        public DelegateCommand CloseCommand { get; set; }

        private void CloseWindow(object parameter)
        {
            IsConfirm = false;
            this.OnClosingRequest();
        }
        #endregion

        #region 绑定属性

        private string _fbfmc;

        public string Fbfmc
        {
            get { return _fbfmc; }
            set
            {
                _fbfmc = value;
                this.RaisePropertyChanged("Fbfmc");
            }
        }
        private string _fbfbm;

        public string Fbfbm
        {
            get { return _fbfbm; }
            set
            {
                _fbfbm = value;
                this.RaisePropertyChanged("Fbfbm");
            }
        }
        private string _fzrxm;

        public string Fzrxm
        {
            get { return _fzrxm; }
            set { _fzrxm = value; this.RaisePropertyChanged("Fzrxm");}
        }

        private string _dcy;

        public string Dcy
        {
            get { return _dcy; }
            set { _dcy = value;this.RaisePropertyChanged("Dcy"); }
        }
        private string _fbfdz;

        public string Fbfdz
        {
            get { return _fbfdz; }
            set { _fbfdz = value;this.RaisePropertyChanged("Fbfdz"); }
        }

        private string _yzbm;

        public string Yzbm
        {
            get { return _yzbm; }
            set { _yzbm = value;this.RaisePropertyChanged("Yzbm"); }
        }
        private string _dcjs;

        public string Dcjs
        {
            get { return _dcjs; }
            set { _dcjs = value; this.RaisePropertyChanged("Dcjs");}
        }
        private string _zjhm;

        public string Zjhm
        {
            get { return _zjhm; }
            set { _zjhm = value; this.RaisePropertyChanged("Zjhm");}
        }
        private string _lxdh;

        public string Lxdh
        {
            get { return _lxdh; }
            set { _lxdh = value; this.RaisePropertyChanged("Lxdh");}
        }

        private DateTime _dcrq;

        public DateTime Dcrq
        {
            get { return _dcrq; }
            set { _dcrq = value; this.RaisePropertyChanged("Dcrq");}
        }

        private List<EntityCode> _zjlxList;

        public List<EntityCode> ZjlxList 
        {
            get { return _zjlxList; }
            set { _zjlxList = value; this.RaisePropertyChanged("ZjlxList"); }
        }

        private EntityCode _fzrxlx;

        public EntityCode Fzrzjlx
        {
            get { return _fzrxlx; }
            set { _fzrxlx = value; this.RaisePropertyChanged("Fzrzjlx");}
        }
        
        #endregion

        #region 操作命令

        public DelegateCommand ConfirmCommand { get; set; }

        private void Save(object parameter)
        {
            if (string.IsNullOrEmpty(Fbfmc) || !Fbfmc.Contains("村民委员会")) return;
            if (string.IsNullOrEmpty(Fbfbm) || Fbfbm.Length != 14) return;
            if (string.IsNullOrEmpty(Fbfdz)) return;
            if (Fzrzjlx == null && string.IsNullOrEmpty(Zjhm)) return;
            this.IsConfirm = true;
            this.OnClosingRequest();

        }
        #endregion
        #region 初始化

        private void InitProperty()
        {
            ZjlxList=new List<EntityCode>()
            {
                new EntityCode(){Code = "1",Entity = "居民身份证"},
                new EntityCode(){Code = "2",Entity = "军官证"},
                new EntityCode(){Code = "3",Entity = "行政、企事业单位机构代码证或法人代码证"},
                new EntityCode(){Code = "4",Entity = "户口簿"},
                new EntityCode(){Code = "5",Entity = "护照"},
                new EntityCode(){Code = "6",Entity = "其他证件"}
            };
            Dcrq=new DateTime(2014,7,10);
            Fbfmc = Fbfbm = Fzrxm = Fbfdz = Yzbm  = Zjhm =Lxdh=Dcy=Dcjs= string.Empty;
            Fzrzjlx = new EntityCode() { Code = "1", Entity = "居民身份证" };
        }
        #endregion

        public FbfInfoViewModel()
        {
            this.CloseCommand=new DelegateCommand();
            this.CloseCommand.ExecuteAction=new Action<object>(CloseWindow);
            this.ConfirmCommand=new DelegateCommand();
            ConfirmCommand.ExecuteAction=new Action<object>(Save);
            InitProperty();
        }
    }
}
