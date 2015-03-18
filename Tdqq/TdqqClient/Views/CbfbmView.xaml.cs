using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using TdqqClient.ViewModels;

namespace TdqqClient.Views
{
    /// <summary>
    /// Interaction logic for CbfbmView.xaml
    /// </summary>
    public partial class CbfbmView : Window
    {
        public CbfbmView(CbfbmViewModel cbfbmVm)
        {
            InitializeComponent();
            this.DataContext = cbfbmVm;
            cbfbmVm.ClosingRequest += (sender, e) => this.Close();
        }
    }
}
