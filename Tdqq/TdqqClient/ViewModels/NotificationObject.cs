using System;
using System.ComponentModel;

namespace TdqqClient.ViewModels
{
    /// <summary>
    /// 对接口INotifyPropertyChanged的实现，并且所有的ViewModel中的类为此基类
    /// </summary>
    public class NotificationObject : INotifyPropertyChanged
    {
        /// <summary>
        /// 属性改变的事件
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        public void RaisePropertyChanged(string propertyName)
        {
            if (this.PropertyChanged != null)
            {
                this.PropertyChanged.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }
        //关闭窗口请求
        public event EventHandler ClosingRequest ;

        protected void OnClosingRequest()
        {
            if (this.ClosingRequest != null)
            {
                this.ClosingRequest(this, EventArgs.Empty);
            }
        }
        //移动窗口请求
        public event EventHandler MoveWindowRequest;

        protected void OnMovingRequest()
        {
            if (this.MoveWindowRequest!=null)
            {
                this.MoveWindowRequest(this, EventArgs.Empty);
            }
        }
    }
}
