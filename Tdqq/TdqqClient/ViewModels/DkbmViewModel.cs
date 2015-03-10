using System;
using TdqqClient.Commands;

namespace TdqqClient.ViewModels
{
    public class DkbmViewModel:NotificationObject
    {

        #region 属性字段
        private string _nsLength;

        public string NsLength
        {
            get { return _nsLength; }
            set
            {
                _nsLength = value;
                this.RaisePropertyChanged("NsLength");
            }
        }

        private double _rowGap;

        public double RowGap
        {
            get { return _rowGap; }
            set
            {
                _rowGap = value;
                this.RaisePropertyChanged("RowGap");
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
        #endregion

        #region 命令属性
        public DelegateCommand MouseLeftButtonDownCommand { get; set; }
        public DelegateCommand ConfirmCommand { get; set; }
        #endregion

        /// <summary>
        /// 属性字段，用来判断是否点击了确定按钮而并非关闭按钮
        /// </summary>
        public bool IsConfirm { get; set; }

        /// <summary>
        /// 关闭按钮
        /// </summary>
        /// <param name="parameter"></param>
        private void CloseWindow(object parameter)
        {
            OnClosingRequest();
            this.IsConfirm = false;
        }

        private void ConfirmButton(object parameter)
        {
            if (RowGap<0||string.IsNullOrEmpty(Fbfbm)||Fbfbm.Length!=14)
            {
                return;
            }
            //请他情况关闭
            OnClosingRequest();
            this.IsConfirm = true;
        }

        public DkbmViewModel(double nsLength)
        {
            this.NsLength = nsLength.ToString("F");
            this.RowGap = 50.0;
            MouseLeftButtonDownCommand=new DelegateCommand();
            ConfirmCommand=new DelegateCommand();
            MouseLeftButtonDownCommand.ExecuteAction=new Action<object>(CloseWindow);
            ConfirmCommand.ExecuteAction=new Action<object>(ConfirmButton);
        }
            
        
        
    }
}
