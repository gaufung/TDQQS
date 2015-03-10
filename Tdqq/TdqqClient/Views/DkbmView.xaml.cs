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
    /// Interaction logic for DkbmView.xaml
    /// </summary>
    public partial class DkbmView : Window
    {
        public DkbmView(DkbmViewModel dkbmVm)
        {
            InitializeComponent();
            this.DataContext = dkbmVm;
            dkbmVm.ClosingRequest += (sender, e) => this.Close();
        }
    }
}
