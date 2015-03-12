using System.Windows;
using TdqqClient.ViewModels;

namespace TdqqClient.Views
{
    /// <summary>
    /// Interaction logic for FieldsView.xaml
    /// </summary>
    public partial class FieldsView : Window
    {
        public FieldsView(FieldsViewModel fieldsVm)
        {
            InitializeComponent();
            this.DataContext = fieldsVm;
            fieldsVm.ClosingRequest += (sender, e) => this.Close();
            fieldsVm.MoveWindowRequest += (sender, e) => this.DragMove();
        }
        
    }
}
