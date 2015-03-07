using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using TdqqClient.ViewModels;

namespace TdqqClient.Views
{
    /// <summary>
    /// Interaction logic for SetDefaultView.xaml
    /// </summary>
    public partial class SetDefaultView : Window
    {
        public SetDefaultView(SetDefaultViewModel setDefaultVm)
        {

            InitializeComponent();
            this.DataContext = setDefaultVm;
            setDefaultVm.ClosingRequest += (sender, e) => this.Close();
           
        }
    }
}
