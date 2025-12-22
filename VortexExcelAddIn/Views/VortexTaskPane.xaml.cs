using System.Windows.Controls;
using VortexExcelAddIn.ViewModels;

namespace VortexExcelAddIn.Views
{
    /// <summary>
    /// Interaction logic for VortexTaskPane.xaml
    /// </summary>
    public partial class VortexTaskPane : UserControl
    {
        public VortexTaskPane()
        {
            InitializeComponent();
            DataContext = new MainViewModel();
        }
    }
}
