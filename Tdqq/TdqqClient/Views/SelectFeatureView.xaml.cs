﻿using System;
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
    /// Interaction logic for SelectFeatureWindow.xaml
    /// </summary>
    public partial class SelectFeatureWindow : Window
    {
        public SelectFeatureWindow(SelectFeatureViewModel selectFeautureWindowVm )
        {
            InitializeComponent();
            this.DataContext = selectFeautureWindowVm;
            selectFeautureWindowVm.ClosingRequest += (sender, e) => this.Close();
        }
    }
}
