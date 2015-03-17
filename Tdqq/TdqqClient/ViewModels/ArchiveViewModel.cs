using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ESRI.ArcGIS.Carto;
using TdqqClient.Commands;
using TdqqClient.Services.Common;

namespace TdqqClient.ViewModels
{
    public class ArchiveViewModel:NotificationObject
    {
        #region 绑定属性
        private bool _isCoverSelected;
        public bool IsCoverSelected
        {
            get { return _isCoverSelected; }
            set
            {
                _isCoverSelected = value;
                this.RaisePropertyChanged("IsCoverSelected");
            }
        }
        private bool _isCbfSelected;

        public bool IsCbfSelected
        {
            get { return _isCbfSelected; }
            set
            {
                _isCbfSelected = value;
                this.RaisePropertyChanged("IsCbfSelected");
            }
        }
        private bool _isDkSelected;

        public bool IsDkSelected
        {
            get { return _isDkSelected; }
            set
            {
                _isDkSelected = value;
                this.RaisePropertyChanged("IsDkSelected");
            }
        }
        private bool _isContractSelected;

        public bool IsContractSelected
        {
            get { return _isContractSelected; }
            set
            {
                _isContractSelected = value;
                this.RaisePropertyChanged("IsContractSelected");
            }
        }

        private bool _isMapSelected;

        public bool IsMapSelected
        {
            get { return _isMapSelected; }
            set
            {
                _isMapSelected = value;
                this.RaisePropertyChanged("IsMapSelected");
            }
        }
        private bool _isStatementSelected;

        public bool IsStatementSelected
        {
            get { return _isStatementSelected; }
            set
            {
                _isStatementSelected = value;
                this.RaisePropertyChanged("IsStatementSelected");
            }
        }
        private bool _isAcceptSelected;

        public bool IsAcceptSelected
        {
            get { return _isAcceptSelected; }
            set { _isAcceptSelected = value; this.RaisePropertyChanged("IsAcceptSelected"); }
        }
        private bool _isGhbSelected;

        public bool IsGhbSelected
        {
            get { return _isGhbSelected; }
            set
            {
                _isGhbSelected = value;
                this.RaisePropertyChanged("IsGhbSelected");
            }
        }
        private bool _isRegisterSelected;

        public bool IsRegisterSelected
        {
            get { return _isRegisterSelected; }
            set
            {
                _isRegisterSelected = value;
                this.RaisePropertyChanged("IsRegisterSelected");
            }
        }
        
        
        
        
        #endregion

        public string CoverFolder { get; set; }
        public string CbfFolder { get; set; }
        public string DkFolder { get; set; }
        public string ContractFolder { get; set; }
        public string MapFolder { get; set; }
        public string StatementFolder { get; set; }
        public string AcceptFolder { get; set; }
        public string GhbFolder { get; set; }
        public string RegisterFolder { get; set; }

        public string ArchiveFolder { get; set; }

        #region 命令属性

        public DelegateCommand CoverCommand { get; set; }
        public DelegateCommand CbfCommand { get; set; }
        public DelegateCommand DkCommand { get; set; }
        public DelegateCommand ContractCommand { get; set; }
        public DelegateCommand MapCommand { get; set; }
        public DelegateCommand StatementCommand { get; set; }

        public DelegateCommand AcceptCommand { get; set; }
        public DelegateCommand GhbCommand { get; set; }

        public DelegateCommand RegisterCommand { get; set; }
        

        #endregion

        #region 操作函数

        private void OpenCoverFolder(object parameter)
        {
            DialogHelper dialogHelper=new DialogHelper();
            CoverFolder = dialogHelper.OpenFolderDialog(false);
            IsCoverSelected = string.IsNullOrEmpty(CoverFolder) ? false : true;           
        }

        private void OpenCbfFolder(object parameter)
        {
            DialogHelper dialogHelper = new DialogHelper();
            CbfFolder = dialogHelper.OpenFolderDialog(false);
            IsCbfSelected = string.IsNullOrEmpty(CbfFolder) ? false : true; 
        }

        private void OpenDkFolder(object parameter)
        {
            DialogHelper dialogHelper = new DialogHelper();
            DkFolder = dialogHelper.OpenFolderDialog(false);
            IsDkSelected = string.IsNullOrEmpty(DkFolder) ? false : true; 
        }

        private void OpenContractFolder(object parameter)
        {
            DialogHelper dialogHelper = new DialogHelper();
            ContractFolder = dialogHelper.OpenFolderDialog(false);
            IsContractSelected = string.IsNullOrEmpty(ContractFolder) ? false : true; 
        }

        private void OpenMapFolder(object parameter)
        {
            DialogHelper dialogHelper = new DialogHelper();
            MapFolder = dialogHelper.OpenFolderDialog(false);
            IsMapSelected = string.IsNullOrEmpty(MapFolder) ? false : true; 
        }

        private void OpenStatementFolder(object parameter)
        {
            DialogHelper dialogHelper = new DialogHelper();
            StatementFolder = dialogHelper.OpenFolderDialog(false);
            IsStatementSelected = string.IsNullOrEmpty(StatementFolder) ? false : true; 
        }

        private void OpenAcceptFolder(object parameter)
        {
            DialogHelper dialogHelper = new DialogHelper();
            AcceptFolder = dialogHelper.OpenFolderDialog(false);
            IsAcceptSelected = string.IsNullOrEmpty(AcceptFolder) ? false : true; 
        }

        private void OpenGhbFolder(object parameter)
        {
            DialogHelper dialogHelper = new DialogHelper();
            GhbFolder = dialogHelper.OpenFolderDialog(false);
            IsGhbSelected = string.IsNullOrEmpty(GhbFolder) ? false : true; 
        }

        private void OpenRegisterFolder(object parameter)
        {
            DialogHelper dialogHelper = new DialogHelper();
            RegisterFolder = dialogHelper.OpenFolderDialog(false);
            IsRegisterSelected = string.IsNullOrEmpty(RegisterFolder) ? false : true; 
        }
        #endregion

        #region 关闭窗口

        public DelegateCommand CloseCommand { get; set; }

        private void CloseWindow(object parameter)
        {
            this.OnClosingRequest();
        }
        public DelegateCommand ArchiveCommand { get; set; }

        private void OpenArchiveFolder(object parameter)
        {
            DialogHelper dialogHelper = new DialogHelper();
            ArchiveFolder = dialogHelper.OpenFolderDialog(true);
            if (string.IsNullOrEmpty(ArchiveFolder)) return;
            this.OnClosingRequest();
        }
        #endregion

        private void InitCommand()
        {
            this.CoverCommand=new DelegateCommand();
            CoverCommand.ExecuteAction=new Action<object>(OpenCoverFolder);
            this.CbfCommand=new DelegateCommand();
            CbfCommand.ExecuteAction=new Action<object>(OpenCbfFolder);
            this.DkCommand=new DelegateCommand();
            DkCommand.ExecuteAction=new Action<object>(OpenDkFolder);
            this.ContractCommand=new DelegateCommand();
            ContractCommand.ExecuteAction=new Action<object>(OpenContractFolder);
            this.MapCommand=new DelegateCommand();
            MapCommand.ExecuteAction=new Action<object>(OpenMapFolder);
            this.StatementCommand=new DelegateCommand();
            StatementCommand.ExecuteAction=new Action<object>(OpenStatementFolder);
            this.AcceptCommand=new DelegateCommand();
            AcceptCommand.ExecuteAction=new Action<object>(OpenAcceptFolder);
            this.GhbCommand=new DelegateCommand();
            GhbCommand.ExecuteAction=new Action<object>(OpenGhbFolder);
            this.RegisterCommand=new DelegateCommand();
            RegisterCommand.ExecuteAction=new Action<object>(OpenRegisterFolder);
            CloseCommand=new DelegateCommand();
            CloseCommand.ExecuteAction=new Action<object>(CloseWindow);
            this.ArchiveCommand=new DelegateCommand();
            ArchiveCommand.ExecuteAction=new Action<object>(OpenArchiveFolder);

        }
        public ArchiveViewModel()
        {
           InitCommand();
        }
    }
}
