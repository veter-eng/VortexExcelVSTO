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
        private int _selectedTabIndex;

        public MainViewModel()
        {
            // Inicializar ViewModels
            ConfigViewModel = new ConfigViewModel();
            QueryViewModel = new QueryViewModel(ConfigViewModel);

            // Começar na aba de configuração
            SelectedTabIndex = 0;

            Services.LoggingService.Info("MainViewModel inicializado");
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
