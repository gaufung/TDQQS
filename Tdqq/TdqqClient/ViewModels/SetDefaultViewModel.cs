using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using NPOI.SS.Formula.Functions;
using TdqqClient.Commands;
using TdqqClient.Models;

namespace TdqqClient.ViewModels
{
    public class SetDefaultViewModel:NotificationObject
    {

        #region 属性字段

        public bool IsConfirm { get; set; }
        private string _zjrxm;

        public string Zjrxm
        {
            get { return _zjrxm; }
            set
            {
                _zjrxm = value;
                this.RaisePropertyChanged("Zjrxm");
            }
        }

        /// <summary>
        /// 土地利用类型
        /// </summary>
        private List<EntityCode> _tdlylxList;
        public List<EntityCode> TdlylxList
        {
            get { return _tdlylxList; }
            set
            {
                _tdlylxList = value;
                this.RaisePropertyChanged("TdlylxList");
            }
        }


        /// <summary>
        /// 承包方经营权取得方式
        /// </summary>
        private List<EntityCode> _cbjyqqdfsList;
        public List<EntityCode> CbjyqqdfsList
        {
            get { return _cbjyqqdfsList; }
            set
            {
                _cbjyqqdfsList = value;
                this.RaisePropertyChanged("CbjyqqdfsList");
            }
        }


        /// <summary>
        /// 所有权性质
        /// </summary>
        private List<EntityCode> _syqxzList;
        public List<EntityCode> SyqxzList
        {
            get { return _syqxzList; }
            set
            {
                _syqxzList = value;
                this.RaisePropertyChanged("SyqxzList");
            }
        }



        private List<EntityCode> _tdytList;
        public List<EntityCode> TdytList
        {
            get { return _tdytList; }
            set
            {
                _tdytList = value;
                this.RaisePropertyChanged("TdytList");
            }
        }

        private EntityCode _tdyt;
        
        public EntityCode Tdyt
        {
            get { return _tdyt; }
            set
            {
                _tdyt = value;
                this.RaisePropertyChanged("Tdyt");
            }
        }
        


        private List<EntityCode> _dldjList;
        public List<EntityCode> DldjList
        {
            get { return _dldjList; }
            set { _dldjList = value; this.RaisePropertyChanged("DldjList"); }
        }

        private EntityCode _dldj;

        public EntityCode Dldj
        {
            get { return _dldj; }
            set
            {
                _dldj = value;
                this.RaisePropertyChanged("Dldj");
            }
        }
        

        private List<EntityCode> _sfjbntList;
        public List<EntityCode> SfjbntList
        {
            get { return _sfjbntList; }
            set { _sfjbntList = value; this.RaisePropertyChanged("SfjbntList"); }
        }

        private EntityCode _sfjbnt;

        public EntityCode Sfjbnt
        {
            get { return _sfjbnt; }
            set
            {
                _sfjbnt = value;
                this.RaisePropertyChanged("Sfjbnt");
            }
        }
        

        private List<EntityCode> _dklbList;
        public List<EntityCode> DklbList
        {
            get { return _dklbList; }
            set { _dklbList = value; this.RaisePropertyChanged("DklbList"); }
        }

        private EntityCode _dklb;

        public EntityCode Dklb
        {
            get { return _dklb; }
            set
            {
                _dklb = value;
                this.RaisePropertyChanged("Dklb");
            }
        }
        

        private EntityCode _tdlylx;

        public EntityCode Tdlylx
        {
            get { return _tdlylx; }
            set { _tdlylx = value;this.RaisePropertyChanged("Tdlylx"); }
        }

        private EntityCode _cbjyqqdfs;

        public EntityCode Cbjyqqdfs
        {
            get { return _cbjyqqdfs; }
            set
            {
                _cbjyqqdfs = value;
                this.RaisePropertyChanged("Cbjyqqdfs");
            }
        }
        private EntityCode _syqxz;

        public EntityCode Syqxz
        {
            get { return _syqxz; }  
            set
            {
                _syqxz = value;
                this.RaisePropertyChanged("Syqxz");
            }
        }
        

     
        
        
        #endregion

        #region 命令属性
        public DelegateCommand MouseLeftButtonDownCommand { get; set; }

        public DelegateCommand ConfirmCommand { get; set; } 
        #endregion

        private void InitComboBox()
        {
            this.TdlylxList=new List<EntityCode>()
            {
                new EntityCode("011","水田"),
                new EntityCode("012","水浇地"),
                new EntityCode("013","旱地"),
            };
            this.CbjyqqdfsList=new List<EntityCode>()
            {
                new EntityCode("100","承包"),
                new EntityCode("110","家庭承包"),
                new EntityCode("120","其他方式承包"),
                new EntityCode("121","招标"),
                new EntityCode("122","拍卖"),
                new EntityCode("123","公开协商"),
                new EntityCode("129","其他方式"),
                new EntityCode("200","转让"),
                new EntityCode("300","互换"),
                new EntityCode("900","其他方式")
            };
            this.SyqxzList=new List<EntityCode>()
            {
                new EntityCode("10","国有土地所有权"),
                new EntityCode("30","集体土地所有权"),
                new EntityCode("31","村民小组"),
                new EntityCode("32","村级集体所有权"),
                new EntityCode("33","乡级集体所有权"),
                new EntityCode("34","其他集体所有权")
            };
            this.TdytList=new List<EntityCode>()
            {
                new EntityCode("1","种植业"),
                new EntityCode("2","林业"),
                new EntityCode("3","畜牧业"),
                new EntityCode("4","渔业"),
                new EntityCode("5","非农业用地")

            };
            this.DldjList=new List<EntityCode>()
            {
                new EntityCode("1","一等地"),
                new EntityCode("2","二等地"),
                new EntityCode("3","三等地"),
                new EntityCode("4","四等地"),
                new EntityCode("5","五等地"),
                new EntityCode("6","六等地"),
                new EntityCode("7","七等地"),
                new EntityCode("8","八等地"),
                new EntityCode("9","九等地"),
                new EntityCode("10","十等地"),
            };
            this.SfjbntList=new List<EntityCode>()
            {
                new EntityCode("1","是"),
                new EntityCode("2","否")
            };
            this.DklbList=new List<EntityCode>()
            {
                new EntityCode("10","承包地块"),
                new EntityCode("21","自留地"),
                new EntityCode("22","机动地"),
                new EntityCode("23","开荒地"),
                new EntityCode("99","其他集体地")
            };
        }

        private void InitCommand()
        {
            this.MouseLeftButtonDownCommand=new DelegateCommand();
            this.ConfirmCommand=new DelegateCommand();
        }
        private void CloseWindow(object parameter)
        {
            OnClosingRequest();
             this.IsConfirm = false;
        }
        /// <summary>
        /// 构造函数
        /// </summary>
        public SetDefaultViewModel()
        {
           InitComboBox();
           IsConfirm = false;
           InitCommand();
           this.MouseLeftButtonDownCommand.ExecuteAction=new Action<object>(CloseWindow);
            this.ConfirmCommand.ExecuteAction=new Action<object>(ConfirmButton);
        }
        
        private void ConfirmButton(object parameter)
        {
            if (Tdlylx==null|| Cbjyqqdfs==null
                || Syqxz==null || Tdyt==null || Dldj==null
                || Sfjbnt==null || Dklb==null)
            {
                return;
            }
            else
            {
                IsConfirm = true;
                OnClosingRequest();
            }
        }

    
        }
        
        
        
}
