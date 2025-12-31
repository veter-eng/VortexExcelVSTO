using System;
using System.Windows;
using VortexExcelAddIn.ViewModels;

namespace VortexExcelAddIn.Views
{
    /// <summary>
    /// Interaction logic for AutoRefreshDialog.xaml
    /// </summary>
    public partial class AutoRefreshDialog : Window
    {
        public AutoRefreshDialog()
        {
            InitializeComponent();

            // Subscrever ao evento RequestClose do ViewModel
            this.Loaded += OnLoaded;
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            if (DataContext is AutoRefreshViewModel viewModel)
            {
                viewModel.RequestClose += OnViewModelRequestClose;
            }
        }

        private void OnViewModelRequestClose(object sender, EventArgs e)
        {
            this.DialogResult = true;
            this.Close();
        }

        /// <summary>
        /// Handler para botão Cancelar/Fechar.
        /// </summary>
        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        protected override void OnClosed(EventArgs e)
        {
            // Desinscrever do evento ao fechar
            if (DataContext is AutoRefreshViewModel viewModel)
            {
                viewModel.RequestClose -= OnViewModelRequestClose;
            }
            base.OnClosed(e);
        }
    }
}
