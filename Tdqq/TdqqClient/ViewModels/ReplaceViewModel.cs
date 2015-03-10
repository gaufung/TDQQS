using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms.PropertyGridInternal;
using TdqqClient.Commands;

namespace TdqqClient.ViewModels
{
    public class ReplaceViewModel:NotificationObject
    {
        private string _originalName;
        public string OriginalName
        {
            get { return _originalName; }
            set
            {
                _originalName = value;
                this.RaisePropertyChanged("OriginalName");
            }
        }


        private string _newName;
        public string NewName
        {
            get { return _newName; }
            set
            {
                _newName = value;
                this.RaisePropertyChanged("NewName");
            }
        }

        public bool IsConfirm { get; set; }

        public DelegateCommand MouseLeftButtonDownCommand { get; set; }
        public DelegateCommand ConfirmCommand { get; set; }

        private void CloseWindow(object parameter)
        {
            OnClosingRequest();
            this.IsConfirm = false;
        }

        public ReplaceViewModel()
        {
            MouseLeftButtonDownCommand=new DelegateCommand();
            ConfirmCommand=new DelegateCommand();
            MouseLeftButtonDownCommand.ExecuteAction=new Action<object>(CloseWindow);
            ConfirmCommand.ExecuteAction=new Action<object>(ConfirmButton);
        }

        private void ConfirmButton(object parameter)
        {
            if (string.IsNullOrEmpty(OriginalName)||string.IsNullOrEmpty(NewName))
            {
                return;
            }
            this.IsConfirm = true;
            OnClosingRequest();
        }

    }
}
