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
    /// Interaction logic for ExportView.xaml
    /// </summary>
    public partial class ExportView : Window
    {
        public ExportView(ExportViewModel exportVm)
        {
            InitializeComponent();
            this.DataContext = exportVm;
            exportVm.ClosingRequest += (sender, e) => this.Close();
        }
    }

    public class CbfbmToShortConvter :IValueConverter
    {

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            string cbfbm = (string) value;
            if (string.IsNullOrEmpty(cbfbm)) return string.Empty;
            return cbfbm.Substring(14);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
