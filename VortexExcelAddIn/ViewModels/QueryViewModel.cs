using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Media;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using VortexExcelAddIn.Models;
using VortexExcelAddIn.Services;

namespace VortexExcelAddIn.ViewModels
{
    /// <summary>
    /// ViewModel para o painel de consulta de dados
    /// Port do QueryPanel.tsx
    /// </summary>
    public partial class QueryViewModel : ViewModelBase
    {
        private readonly ConfigViewModel _configViewModel;

        /// <summary>
        /// Evento disparado quando dados são exportados para o Excel
        /// </summary>
        public event EventHandler DataExported;

        #region Observable Properties

        // Campos de texto para IDs (separados por vírgula)
        [ObservableProperty]
        private string _coletorIds;

        [ObservableProperty]
        private string _gatewayIds;

        [ObservableProperty]
        private string _equipmentIds;

        [ObservableProperty]
        private string _tagIds;

        // Parâmetros de tempo
        [ObservableProperty]
        private DateTime _startDate;

        [ObservableProperty]
        private DateTime _endDate;

        [ObservableProperty]
        private int _limit;

        // Estados
        [ObservableProperty]
        private bool _isQuerying;

        // Resultados
        [ObservableProperty]
        private ObservableCollection<VortexDataPoint> _results;

        [ObservableProperty]
        private ObservableCollection<VortexDataPoint> _previewResults;

        // Mensagens
        [ObservableProperty]
        private string _statusMessage;

        [ObservableProperty]
        private Brush _statusMessageColor;

        // Debug info
        [ObservableProperty]
        private string _lastQueryExecuted;

        [ObservableProperty]
        private string _debugInfo;

        #endregion

        #region Computed Properties

        /// <summary>
        /// Indica se o tipo de servidor selecionado é VortexIO (API com dados agregados)
        /// </summary>
        public bool IsVortexIO => _configViewModel?.SelectedDatabaseType == Domain.Models.DatabaseType.VortexAPI;

        /// <summary>
        /// Indica se o tipo de servidor selecionado é Vortex Historian (API com dados brutos)
        /// </summary>
        public bool IsVortexHistorian => !IsVortexIO;

        #endregion

        public QueryViewModel(ConfigViewModel configViewModel)
        {
            _configViewModel = configViewModel ?? throw new ArgumentNullException(nameof(configViewModel));

            // Subscrever mudanças no ConfigViewModel para atualizar IsVortexIO/IsVortexHistorian
            _configViewModel.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == nameof(ConfigViewModel.SelectedDatabaseType))
                {
                    OnPropertyChanged(nameof(IsVortexIO));
                    OnPropertyChanged(nameof(IsVortexHistorian));
                }
            };

            // Inicializar coleções de resultados
            Results = new ObservableCollection<VortexDataPoint>();
            PreviewResults = new ObservableCollection<VortexDataPoint>();

            // Valores padrão
            StartDate = DateTime.Now.AddHours(-24);
            EndDate = DateTime.Now;
            Limit = 1000;

            // Inicializar campos vazios (pegar todos)
            ColetorIds = string.Empty;
            GatewayIds = string.Empty;
            EquipmentIds = string.Empty;
            TagIds = string.Empty;

            StatusMessageColor = Brushes.Gray;
            StatusMessage = "Digite os filtros ou deixe vazio para buscar todos os dados";

            LoggingService.Info("QueryViewModel inicializado");
        }


        #region Commands

        /// <summary>
        /// Comando para executar consulta
        /// </summary>
        [RelayCommand]
        private async Task QueryAsync()
        {
            IsQuerying = true;
            StatusMessage = "Consultando dados...";
            StatusMessageColor = Brushes.Gray;
            Results.Clear();
            PreviewResults.Clear();

            try
            {
                // Usar nova arquitetura com IDataSourceConnection (SOLID: DIP)
                var connection = _configViewModel.GetConnection();

                if (connection == null)
                {
                    StatusMessage = "Configure a Conexão com o banco de dados na aba 'Configuração' antes de consultar dados";
                    StatusMessageColor = Brushes.Orange;
                    LoggingService.Warn("Tentativa de consulta sem configuração válida");
                    return;
                }

                // Log para debug: qual tipo de banco está sendo usado
                var connInfo = connection.GetConnectionInfo();
                LoggingService.Info($"[QUERY DEBUG] Executando query com DatabaseType: {connInfo.DatabaseType}, Connection: {connInfo.DatabaseName}");

                // Criar parâmetros de consulta
                var queryParams = new QueryParams
                {
                    ColetorId = ColetorIds,
                    GatewayId = GatewayIds,
                    EquipmentId = EquipmentIds,
                    TagId = TagIds,
                    StartTime = StartDate,
                    EndTime = EndDate,
                    Limit = Limit
                };

                // Executar query usando interface (funciona com qualquer banco)
                var data = await connection.QueryDataAsync(queryParams);

                // Capturar informações de debug
                var connectionInfo = connection.GetConnectionInfo();
                LastQueryExecuted = $"Tipo: {connectionInfo.DatabaseType}, Banco: {connectionInfo.DatabaseName}";
                DebugInfo = $"Conexão: {connectionInfo}\n\nPrimeiros 3 registros retornados:\n{string.Join("\n", data.Take(3).Select((d, idx) => $"{idx + 1}. Time={d.Time:dd/MM/yyyy HH:mm:ss}, Valor=[{d.Valor}], Coletor=[{d.ColetorId}], Gateway=[{d.GatewayId}], Equip=[{d.EquipmentId}], Tag=[{d.TagId}]"))}";

                // Atualizar resultados
                Results.Clear();
                foreach (var point in data)
                {
                    Results.Add(point);
                }

                // Preview (primeiros 10)
                PreviewResults.Clear();
                foreach (var point in data.Take(10))
                {
                    PreviewResults.Add(point);
                }

                StatusMessage = $"Consulta Concluída: {Results.Count:N0} Registros Válidos Encontrados!";
                StatusMessageColor = Brushes.Green;
                LoggingService.Info($"Consulta retornou {Results.Count:N0} registros válidos");
            }
            catch (Exception ex)
            {
                StatusMessage = $"Erro na consulta: {ex.Message}";
                StatusMessageColor = Brushes.Red;
                LoggingService.Error("Erro ao executar consulta", ex);
            }
            finally
            {
                IsQuerying = false;
            }
        }

        /// <summary>
        /// Comando para exportar para Excel
        /// </summary>
        [RelayCommand]
        private void ExportToSheet()
        {
            if (Results == null || Results.Count == 0)
            {
                StatusMessage = "Nenhum dado para exportar";
                StatusMessageColor = Brushes.Orange;
                return;
            }

            try
            {
                var dataList = Results.ToList();
                ExcelService.ExportToSheet(dataList, $"VortexData_{DateTime.Now:yyyyMMdd_HHmmss}", _configViewModel.SelectedDatabaseType);

                StatusMessage = $"Dados Exportados para o Excel: {Results.Count:N0} Registros";
                StatusMessageColor = Brushes.Green;
                LoggingService.Info($"Dados Exportados para o Excel: {Results.Count:N0} Registros");

                // Notificar que dados foram exportados
                DataExported?.Invoke(this, EventArgs.Empty);
            }
            catch (Exception ex)
            {
                StatusMessage = $"Erro ao exportar: {ex.Message}";
                StatusMessageColor = Brushes.Red;
                LoggingService.Error("Erro ao exportar para Excel", ex);
            }
        }

        /// <summary>
        /// Comando para baixar CSV
        /// </summary>
        [RelayCommand]
        private void DownloadCsv()
        {
            if (Results == null || Results.Count == 0)
            {
                StatusMessage = "Nenhum dado para baixar";
                StatusMessageColor = Brushes.Orange;
                return;
            }

            try
            {
                var dataList = Results.ToList();
                ExcelService.DownloadCsv(dataList, $"VortexData_{DateTime.Now:yyyyMMdd_HHmmss}.csv");

                LoggingService.Info($"CSV baixado: {Results.Count} registros");
            }
            catch (Exception ex)
            {
                StatusMessage = $"Erro ao baixar CSV: {ex.Message}";
                StatusMessageColor = Brushes.Red;
                LoggingService.Error("Erro ao baixar CSV", ex);
            }
        }

        #endregion
    }
}
