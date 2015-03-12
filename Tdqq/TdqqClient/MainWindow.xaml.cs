using System.Windows;
using System.Windows.Forms;
using ESRI.ArcGIS.Controls;
using TdqqClient.ViewModels;

namespace TdqqClient
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private AxMapControl _mainMapControl;
        public MainWindow()
        {
            InitializeComponent();
            CreateEngineControls();
            var mainVm = new MainWindowViewModel(_mainMapControl);
            this.DataContext = mainVm;
            mainVm.ParentWindow = this;

        }
        private void CreateEngineControls()
        {
            _mainMapControl = new AxMapControl();
            _mainMapControl.Dock = DockStyle.None;
            _mainMapControl.BackColor = System.Drawing.Color.AliceBlue;
            MainFormsHost.Child = _mainMapControl;
        }
    }
}
