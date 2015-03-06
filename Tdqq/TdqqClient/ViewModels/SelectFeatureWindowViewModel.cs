using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using TdqqClient.Commands;
using TdqqClient.Services.AE;

namespace TdqqClient.ViewModels
{
    public class SelectFeatureWindowViewModel:NotificationObject
    {
        #region 属性

        private string _caption;

        public string Caption
        {
            get { return _caption; }
            set
            {
                _caption = value;
            }
        }

        private string _selectFeature;

        public string SelectFeature
        {
            get { return _selectFeature; }
            set
            {
                _selectFeature = value;
                this.RaisePropertyChanged("SelectFeature");
            }
        }
        private List<string> _listFeatures;

        public List<string> ListFeautrues
        {
            get { return _listFeatures; }
            set
            {
                _listFeatures = value;
                this.RaisePropertyChanged("ListFeautrues");
            }
        } 
        #endregion

        private readonly string _personDatabse;
        public DelegateCommand MouseLeftButtonDownCommand { get; set; }

        public DelegateCommand ConfirmCommand { get; set; }

        private void CloseWindow(object parameter)
        {
            OnClosingRequest();
        }
        private void ConfirmButton(object parameter)
        {
            if (string.IsNullOrEmpty(SelectFeature))
            {
                return;
            }
            else
            {
                OnClosingRequest();  
            }           
        }
        private void InitCommand()
        {
            MouseLeftButtonDownCommand=new DelegateCommand();
            ConfirmCommand=new DelegateCommand();
            MouseLeftButtonDownCommand.ExecuteAction=new Action<object>(CloseWindow);
            ConfirmCommand.ExecuteAction=new Action<object>(ConfirmButton);
        }
        public SelectFeatureWindowViewModel(string personDatabase)
        {
            _personDatabse = personDatabase;
            this.Caption = "请选择地块要素类";
            ListFeautrues = AeHelper.GetAllFeautureClass(_personDatabse);
            InitCommand();
        }
    }
}
