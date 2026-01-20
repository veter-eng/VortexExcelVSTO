using System;
using System.Windows;
using VortexExcelAddIn.ViewModels;

namespace VortexExcelAddIn.Views
{
    /// <summary>
    /// Lógica de interação para TempoDialog.xaml
    /// Diálogo para configuração de agregação temporal.
    /// </summary>
    public partial class TempoDialog : Window
    {
        public TempoDialog()
        {
            InitializeComponent();

            // Subscrever ao evento RequestClose do ViewModel
            Loaded += (s, e) =>
            {
                if (DataContext is TempoViewModel vm)
                {
                    vm.RequestClose += OnViewModelRequestClose;

                    // Inicializar o PasswordBox com o token existente
                    if (!string.IsNullOrEmpty(vm.Token))
                    {
                        TokenPasswordBox.Password = vm.Token;
                    }
                }
            };

            Unloaded += (s, e) =>
            {
                if (DataContext is TempoViewModel vm)
                {
                    vm.RequestClose -= OnViewModelRequestClose;
                }
            };
        }

        private void OnViewModelRequestClose(object sender, EventArgs e)
        {
            this.DialogResult = true;
            this.Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        private void TokenPasswordBox_PasswordChanged(object sender, RoutedEventArgs e)
        {
            if (DataContext is TempoViewModel vm)
            {
                vm.Token = TokenPasswordBox.Password;
            }
        }

        private void InfoButton_Click(object sender, RoutedEventArgs e)
        {
            var message = "📊 Como Funciona a Agregação Temporal\n\n" +
                          "• Múltiplas seleções:\n" +
                          "  Você pode selecionar vários tipos de agregação e janelas de tempo\n\n" +
                          "• Resultado:\n" +
                          "  Dados para TODAS as combinações selecionadas serão retornados\n\n" +
                          "• Exemplo:\n" +
                          "  Média + Máximo com 5min + 60min = 4 conjuntos de dados\n\n" +
                          "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n" +
                          "🔄 Vortex Historian:\n" +
                          "  Aplica agregação em tempo real aos dados brutos usando Flux queries\n\n" +
                          "🔍 VortexIO:\n" +
                          "  Filtra dados já pré-agregados pelo Airflow (não re-agrega)";

            MessageBox.Show(message, "ℹ️ Informações sobre Agregação", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}
