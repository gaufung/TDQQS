using System;
using System.Windows;
using System.Windows.Threading;

namespace TdqqClient.Views
{
    /// <summary>
    /// Interaction logic for Wait.xaml
    /// </summary>
    public partial class Wait : Window
    {
        public Wait()
        {
            InitializeComponent();
            InitControls();
        }
        private void InitControls()
        {
            this.labelCaption.Content = string.Empty;
            this.progressBar.Minimum = 0;
            this.progressBar.Maximum = 100;
            this.progressBar.Value = 0;
        }
        /// <summary>
        /// 设置等待窗体的标题
        /// </summary>
        /// <param name="caption">标题名称</param>
        public void SetWaitCaption(string caption)
        {
            this.labelCaption.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(() =>
            {
                this.labelCaption.Content = caption;
            }));
        }
        /// <summary>
        /// 设置等待窗体进度条
        /// </summary>
        /// <param name="progress"></param>
        public void SetProgress(double progress)
        {
            const double scaleFactor = 100;
            this.progressBar.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(() =>
            {
                this.progressBar.Value = progress * scaleFactor;
            }));
        }
        /// <summary>
        /// 关闭等待窗口
        /// </summary>
        public void CloseWait()
        {
            this.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(() => Close()));
        }
    }
}
