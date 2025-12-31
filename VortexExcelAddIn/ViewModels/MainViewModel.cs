using System;
using CommunityToolkit.Mvvm.ComponentModel;

namespace VortexExcelAddIn.ViewModels
{
    /// <summary>
    /// ViewModel principal que agrega Config e Query ViewModels
    /// </summary>
    public partial class MainViewModel : ViewModelBase
    {
        [ObservableProperty]
        private ConfigViewModel _configViewModel;

        [ObservableProperty]
        private QueryViewModel _queryViewModel;

        [ObservableProperty]
        private AutoRefreshViewModel _autoRefreshViewModel;

        [ObservableProperty]
        private int _selectedTabIndex;

        [ObservableProperty]
        private bool _hasDataBeenExported;

        public MainViewModel()
        {
            // Inicializar ViewModels
            ConfigViewModel = new ConfigViewModel();
            QueryViewModel = new QueryViewModel(ConfigViewModel);

            // Subscrever ao evento de exportação de dados
            QueryViewModel.DataExported += OnDataExported;

            // Subscrever ao evento de conexão bem-sucedida
            ConfigViewModel.ConnectionSavedSuccessfully += OnConnectionSavedSuccessfully;

            // Começar na aba de configuração
            SelectedTabIndex = 0;
            HasDataBeenExported = false;

            Services.LoggingService.Info("MainViewModel inicializado");
        }

        /// <summary>
        /// Evento disparado quando dados são exportados
        /// </summary>
        public event EventHandler DataExportedToExcel;

        private void OnDataExported(object sender, EventArgs e)
        {
            HasDataBeenExported = true;
            Services.LoggingService.Debug("Dados foram exportados - Refresh habilitado");

            // Notificar ThisAddIn via evento
            DataExportedToExcel?.Invoke(this, EventArgs.Empty);
        }

        /// <summary>
        /// Handler para evento de conexão salva com sucesso.
        /// Navega automaticamente para a aba de Query.
        /// </summary>
        private void OnConnectionSavedSuccessfully(object sender, EventArgs e)
        {
            SelectedTabIndex = 1; // Aba de Query
            Services.LoggingService.Debug("Navegando para aba de Query após conexão bem-sucedida");
        }

        /// <summary>
        /// Cleanup
        /// </summary>
        public void Dispose()
        {
            ConfigViewModel?.Dispose();
            Services.LoggingService.Debug("MainViewModel disposed");
        }
    }
}
