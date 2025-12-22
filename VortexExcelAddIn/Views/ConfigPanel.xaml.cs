using System.Windows.Controls;
using VortexExcelAddIn.ViewModels;

namespace VortexExcelAddIn.Views
{
    /// <summary>
    /// Interaction logic for ConfigPanel.xaml
    /// </summary>
    public partial class ConfigPanel : UserControl
    {
        public ConfigPanel()
        {
            InitializeComponent();

            // Sincronizar PasswordBox com ViewModel (PasswordBox não suporta binding direto)
            TokenPasswordBox.PasswordChanged += (s, e) =>
            {
                if (DataContext is ConfigViewModel vm)
                {
                    vm.Token = TokenPasswordBox.Password;
                }
            };

            DataContextChanged += (s, e) =>
            {
                if (DataContext is ConfigViewModel vm && !string.IsNullOrEmpty(vm.Token))
                {
                    TokenPasswordBox.Password = vm.Token;
                }
            };
        }
    }
}
