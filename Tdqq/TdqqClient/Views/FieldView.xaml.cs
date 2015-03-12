using System.Windows;
using TdqqClient.ViewModels;

namespace TdqqClient.Views
{
    /// <summary>
    /// Interaction logic for FieldView.xaml
    /// </summary>
    public partial class FieldView : Window
    {
        public FieldView()
        {
            InitializeComponent();
           
        }
        public void SetDataContext(FieldViewModel fieldvm)
        {
            this.DataContext = fieldvm;
            fieldvm.ClosingRequest += (sender, e) => this.Close();
            fieldvm.MoveWindowRequest += (sener, e) => this.DragMove();
        }
        
    }
}
