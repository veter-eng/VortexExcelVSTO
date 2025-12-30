using System.Windows;
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
        }

        /// <summary>
        /// Handler para mudanças na senha do PasswordBox.
        /// PasswordBox não suporta binding direto por motivos de segurança,
        /// então usamos este handler para atualizar o ViewModel.
        /// </summary>
        private void PasswordBox_OnPasswordChanged(object sender, RoutedEventArgs e)
        {
            if (DataContext is ConfigViewModel viewModel && sender is PasswordBox passwordBox)
            {
                viewModel.Password = passwordBox.Password;
            }
        }
    }
}
