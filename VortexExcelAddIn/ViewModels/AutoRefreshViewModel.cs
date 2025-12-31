using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Media;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using VortexExcelAddIn.Domain.Interfaces;
using VortexExcelAddIn.Models;
using VortexExcelAddIn.Services;
using Excel = Microsoft.Office.Interop.Excel;

namespace VortexExcelAddIn.ViewModels
{
    /// <summary>
    /// ViewModel para configuração de auto-refresh e exibição de status.
    /// Segue padrão MVVM usando CommunityToolkit.Mvvm.
    /// </summary>
    public partial class AutoRefreshViewModel : ViewModelBase
    {
        private readonly IAutoRefreshService _autoRefreshService;
        private readonly QueryViewModel _queryViewModel;

        /// <summary>
        /// Evento disparado quando o diálogo deve ser fechado.
        /// </summary>
        public event EventHandler RequestClose;

        #region Observable Properties

        [ObservableProperty]
        private bool _isEnabled;

        [ObservableProperty]
        private int _intervalMinutes;

        [ObservableProperty]
        private string _targetSheetName;

        [ObservableProperty]
        private ObservableCollection<string> _availableSheets;

        [ObservableProperty]
        private bool _isActive;

        [ObservableProperty]
        private DateTime? _nextRefreshTime;

        [ObservableProperty]
        private string _statusMessage;

        [ObservableProperty]
        private Brush _statusColor;

        [ObservableProperty]
        private DateTime? _lastRefreshTime;

        [ObservableProperty]
        private int _lastRecordCount;

        #endregion

        /// <summary>
        /// Inicializa uma nova instância de AutoRefreshViewModel.
        /// </summary>
        public AutoRefreshViewModel(
            IAutoRefreshService autoRefreshService,
            QueryViewModel queryViewModel)
        {
            _autoRefreshService = autoRefreshService ?? throw new ArgumentNullException(nameof(autoRefreshService));
            _queryViewModel = queryViewModel ?? throw new ArgumentNullException(nameof(queryViewModel));

            // Valores padrão
            IntervalMinutes = 5;
            TargetSheetName = string.Empty;
            AvailableSheets = new ObservableCollection<string>();
            StatusColor = Brushes.Gray;
            StatusMessage = "Refresh automático não configurado";

            // Subscrever eventos do serviço
            _autoRefreshService.RefreshStarted += OnRefreshStarted;
            _autoRefreshService.RefreshCompleted += OnRefreshCompleted;
            _autoRefreshService.RefreshFailed += OnRefreshFailed;
            _autoRefreshService.NextRefreshTimeChanged += OnNextRefreshTimeChanged;

            // Carregar configurações atuais
            LoadCurrentSettings();
            RefreshAvailableSheets();
        }

        #region Commands

        /// <summary>
        /// Comando para iniciar auto-refresh.
        /// </summary>
        [RelayCommand]
        private void StartAutoRefresh()
        {
            try
            {
                // Validar intervalo
                if (IntervalMinutes < 1 || IntervalMinutes > 60)
                {
                    StatusMessage = "Intervalo deve estar entre 1 e 60 minutos";
                    StatusColor = Brushes.Red;
                    return;
                }

                // Criar configurações a partir do estado atual do ViewModel
                string targetSheet;
                if (TargetSheetName == "(Planilha Atual)")
                {
                    // Null indica que deve usar a planilha ativa no momento do refresh
                    targetSheet = null;
                }
                else if (TargetSheetName == "(Criar nova planilha a cada atualização)")
                {
                    // String vazia indica criar nova planilha
                    targetSheet = string.Empty;
                }
                else
                {
                    // Nome específico de planilha
                    targetSheet = TargetSheetName;
                }

                var settings = new AutoRefreshSettings
                {
                    IsEnabled = true,
                    IntervalMinutes = IntervalMinutes,
                    TargetSheetName = targetSheet,
                    ResultLimit = 1000,
                    QueryParameters = CaptureQueryParameters()
                };

                // Salvar configurações
                AutoRefreshConfigService.SaveSettings(settings);

                // Iniciar serviço
                _autoRefreshService.Start(settings);

                IsActive = true;
                IsEnabled = true;
                StatusMessage = $"Refresh iniciado: a cada {IntervalMinutes} minutos";
                StatusColor = Brushes.Green;

                LoggingService.Info("Auto-refresh iniciado pelo usuário");

                // Fechar o diálogo após sucesso
                RequestClose?.Invoke(this, EventArgs.Empty);
            }
            catch (Exception ex)
            {
                StatusMessage = $"Falha ao iniciar: {ex.Message}";
                StatusColor = Brushes.Red;
                LoggingService.Error("Erro ao iniciar auto-refresh", ex);
            }
        }

        /// <summary>
        /// Comando para parar auto-refresh.
        /// </summary>
        [RelayCommand]
        private void StopAutoRefresh()
        {
            try
            {
                // Atualizar configurações para desabilitado
                var settings = AutoRefreshConfigService.LoadSettings();
                if (settings != null)
                {
                    settings.IsEnabled = false;
                    AutoRefreshConfigService.SaveSettings(settings);
                }

                // Parar serviço
                _autoRefreshService.Stop();

                IsActive = false;
                IsEnabled = false;
                NextRefreshTime = null;
                StatusMessage = "Refresh parado";
                StatusColor = Brushes.Gray;

                LoggingService.Info("Auto-refresh parado pelo usuário");
            }
            catch (Exception ex)
            {
                StatusMessage = $"Falha ao parar: {ex.Message}";
                StatusColor = Brushes.Red;
                LoggingService.Error("Erro ao parar auto-refresh", ex);
            }
        }

        /// <summary>
        /// Comando para atualizar agora (manual).
        /// </summary>
        [RelayCommand]
        private async Task RefreshNow()
        {
            try
            {
                StatusMessage = "Atualizando agora...";
                StatusColor = Brushes.Blue;

                await _autoRefreshService.RefreshNowAsync();
            }
            catch (Exception ex)
            {
                StatusMessage = $"Atualização falhou: {ex.Message}";
                StatusColor = Brushes.Red;
                LoggingService.Error("Atualização manual falhou", ex);
            }
        }

        /// <summary>
        /// Comando para atualizar lista de planilhas disponíveis.
        /// </summary>
        [RelayCommand]
        private void RefreshSheetList()
        {
            RefreshAvailableSheets();
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Captura os parâmetros de consulta atuais do QueryViewModel.
        /// </summary>
        private QueryParams CaptureQueryParameters()
        {
            return new QueryParams
            {
                ColetorId = _queryViewModel.ColetorIds,
                GatewayId = _queryViewModel.GatewayIds,
                EquipmentId = _queryViewModel.EquipmentIds,
                TagId = _queryViewModel.TagIds,
                StartTime = _queryViewModel.StartDate,
                EndTime = _queryViewModel.EndDate,
                Limit = 1000 // Fixo conforme requisitos
            };
        }

        /// <summary>
        /// Atualiza lista de planilhas disponíveis.
        /// </summary>
        private void RefreshAvailableSheets()
        {
            try
            {
                AvailableSheets.Clear();
                AvailableSheets.Add("(Planilha Atual)");
                AvailableSheets.Add("(Criar nova planilha a cada atualização)");

                var workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                if (workbook != null)
                {
                    foreach (Excel.Worksheet sheet in workbook.Worksheets)
                    {
                        AvailableSheets.Add(sheet.Name);
                    }
                }

                // Selecionar "(Planilha Atual)" como padrão
                if (string.IsNullOrEmpty(TargetSheetName) && AvailableSheets.Count > 0)
                {
                    TargetSheetName = AvailableSheets[0]; // "(Planilha Atual)"
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao carregar planilhas disponíveis", ex);
            }
        }

        /// <summary>
        /// Carrega configurações atuais do serviço.
        /// </summary>
        private void LoadCurrentSettings()
        {
            var settings = _autoRefreshService.Settings;
            if (settings != null)
            {
                // Carregar configurações salvas, mas NÃO usar IsEnabled
                // pois isso é apenas configuração, não estado de execução
                IntervalMinutes = settings.IntervalMinutes;

                // Converter de volta para o texto da UI
                if (settings.TargetSheetName == null)
                {
                    TargetSheetName = "(Planilha Atual)";
                }
                else if (string.IsNullOrEmpty(settings.TargetSheetName))
                {
                    TargetSheetName = "(Criar nova planilha a cada atualização)";
                }
                else
                {
                    TargetSheetName = settings.TargetSheetName;
                }

                LastRefreshTime = settings.LastRefreshTime;
            }

            // Sempre pegar o estado atual do serviço (não das configurações)
            IsActive = _autoRefreshService.IsActive;
            NextRefreshTime = _autoRefreshService.NextRefreshTime;
            IsEnabled = IsActive; // IsEnabled reflete o estado atual, não a configuração salva

            if (IsActive)
            {
                StatusMessage = $"Refresh ativo: a cada {IntervalMinutes} minutos";
                StatusColor = Brushes.Green;
            }
        }

        #endregion

        #region Event Handlers

        /// <summary>
        /// Handler para evento de início de atualização.
        /// </summary>
        private void OnRefreshStarted(object sender, EventArgs e)
        {
            StatusMessage = "Atualizando dados...";
            StatusColor = Brushes.Blue;
        }

        /// <summary>
        /// Handler para evento de atualização concluída.
        /// </summary>
        private void OnRefreshCompleted(object sender, RefreshCompletedEventArgs e)
        {
            LastRefreshTime = e.RefreshTime;
            LastRecordCount = e.RecordsUpdated;
            StatusMessage = $"Última atualização: {e.RecordsUpdated} registros em {e.Duration.TotalSeconds:F1}s";
            StatusColor = Brushes.Green;
        }

        /// <summary>
        /// Handler para evento de falha na atualização.
        /// </summary>
        private void OnRefreshFailed(object sender, RefreshErrorEventArgs e)
        {
            StatusMessage = $"Atualização falhou: {e.Error.Message}";
            StatusColor = Brushes.Red;
        }

        /// <summary>
        /// Handler para evento de mudança no próximo horário de atualização.
        /// </summary>
        private void OnNextRefreshTimeChanged(object sender, EventArgs e)
        {
            NextRefreshTime = _autoRefreshService.NextRefreshTime;
        }

        #endregion
    }
}
