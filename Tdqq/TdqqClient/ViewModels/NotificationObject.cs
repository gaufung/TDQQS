using System.ComponentModel;
namespace TdqqClient.ViewModels
{
    /// <summary>
    /// 对接口INotifyPropertyChanged的实现，并且所有的ViewModel中的类为此基类
    /// </summary>
    class NotificationObject : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        public void RaisePropertyChanged(string propertyName)
        {
            if (this.PropertyChanged != null)
            {
                this.PropertyChanged.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
}
