using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using TdqqClient.ViewModels;

namespace TdqqClient.Views
{
    /// <summary>
    /// Interaction logic for ArchiveView.xaml
    /// </summary>
    public partial class ArchiveView : Window
    {
        public ArchiveView(ArchiveViewModel archiveVm)
        {
            InitializeComponent();
            this.DataContext = archiveVm;
            archiveVm.ClosingRequest += (sender, e) => this.Close();
        }
    }

    public class VisibilityToBoolConverter : IValueConverter
    {

        public object Convert(object value, System.Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
           
            bool isSelected = (bool)value;
            if (isSelected)
            {
                return Visibility.Visible;
            }
            else
            {
                return Visibility.Hidden;
            }
        }

        public object ConvertBack(object value, System.Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            Visibility visiblity = (Visibility)value;
            switch (visiblity)
            {
                case Visibility.Visible:
                    return true;
                case Visibility.Hidden:
                    return false;
                default:
                    return false;
            }
        }
    }
}
