using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using TdqqClient.ViewModels;

namespace TdqqClient.Views
{
    /// <summary>
    /// Interaction logic for FarmersView.xaml
    /// </summary>
    public partial class FarmersView : Window
    {
        public FarmersView(FarmersViewModel farmersVm)
        {
            InitializeComponent();
            this.DataContext = farmersVm;
            farmersVm.ClosingRequest += (sender, e) => this.Close();
            farmersVm.MoveWindowRequest += (sender, e) => this.DragMove();
        }
    }
}
