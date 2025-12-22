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

        #region Observable Properties

        // Listas de filtros (cascata)
        [ObservableProperty]
        private ObservableCollection<string> _coletores;

        [ObservableProperty]
        private ObservableCollection<string> _gateways;

        [ObservableProperty]
        private ObservableCollection<string> _equipments;

        [ObservableProperty]
        private ObservableCollection<string> _tags;

        // Seleções
        [ObservableProperty]
        private string _selectedColetor;

        [ObservableProperty]
        private string _selectedGateway;

        [ObservableProperty]
        private string _selectedEquipment;

        [ObservableProperty]
        private string _selectedTag;

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

        [ObservableProperty]
        private bool _isLoadingFilters;

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

        #endregion

        public QueryViewModel(ConfigViewModel configViewModel)
        {
            _configViewModel = configViewModel ?? throw new ArgumentNullException(nameof(configViewModel));

            // Inicializar coleções
            Coletores = new ObservableCollection<string>();
            Gateways = new ObservableCollection<string>();
            Equipments = new ObservableCollection<string>();
            Tags = new ObservableCollection<string>();
            Results = new ObservableCollection<VortexDataPoint>();
            PreviewResults = new ObservableCollection<VortexDataPoint>();

            // Valores padrão
            StartDate = DateTime.Now.AddHours(-24);
            EndDate = DateTime.Now;
            Limit = 1000;

            StatusMessageColor = Brushes.Gray;
            StatusMessage = "Selecione os filtros e clique em Consultar";

            // Carregar coletores inicialmente (se conectado)
            if (_configViewModel.IsConnected)
            {
                _ = LoadColetoresAsync();
            }
        }

        #region Property Changed Handlers (Cascata)

        /// <summary>
        /// Quando o coletor muda, carrega gateways
        /// </summary>
        partial void OnSelectedColetorChanged(string value)
        {
            if (!string.IsNullOrEmpty(value))
            {
                _ = LoadGatewaysAsync(value);
            }
            else
            {
                Gateways.Clear();
                Equipments.Clear();
                Tags.Clear();
                SelectedGateway = null;
            }
        }

        /// <summary>
        /// Quando o gateway muda, carrega equipamentos
        /// </summary>
        partial void OnSelectedGatewayChanged(string value)
        {
            if (!string.IsNullOrEmpty(value) && !string.IsNullOrEmpty(SelectedColetor))
            {
                _ = LoadEquipmentsAsync(SelectedColetor, value);
            }
            else
            {
                Equipments.Clear();
                Tags.Clear();
                SelectedEquipment = null;
            }
        }

        /// <summary>
        /// Quando o equipamento muda, carrega tags
        /// </summary>
        partial void OnSelectedEquipmentChanged(string value)
        {
            if (!string.IsNullOrEmpty(value) && !string.IsNullOrEmpty(SelectedColetor) && !string.IsNullOrEmpty(SelectedGateway))
            {
                _ = LoadTagsAsync(SelectedColetor, SelectedGateway, value);
            }
            else
            {
                Tags.Clear();
                SelectedTag = null;
            }
        }

        #endregion

        #region Load Filters Methods

        /// <summary>
        /// Carrega lista de coletores
        /// </summary>
        private async Task LoadColetoresAsync()
        {
            try
            {
                IsLoadingFilters = true;
                var service = _configViewModel.GetInfluxDbService();

                if (service == null)
                {
                    StatusMessage = "Configure a conexão com o InfluxDB primeiro";
                    StatusMessageColor = Brushes.Orange;
                    return;
                }

                var coletores = await service.GetAvailableCollectorsAsync();
                Coletores.Clear();

                foreach (var coletor in coletores.OrderBy(c => c))
                {
                    Coletores.Add(coletor);
                }

                LoggingService.Debug($"Carregados {Coletores.Count} coletores");
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao carregar coletores", ex);
                StatusMessage = $"Erro ao carregar coletores: {ex.Message}";
                StatusMessageColor = Brushes.Red;
            }
            finally
            {
                IsLoadingFilters = false;
            }
        }

        /// <summary>
        /// Carrega gateways para um coletor
        /// </summary>
        private async Task LoadGatewaysAsync(string coletorId)
        {
            try
            {
                IsLoadingFilters = true;
                Gateways.Clear();
                Equipments.Clear();
                Tags.Clear();

                var service = _configViewModel.GetInfluxDbService();
                if (service == null) return;

                var gateways = await service.GetAvailableGatewaysAsync(coletorId);

                foreach (var gateway in gateways.OrderBy(g => g))
                {
                    Gateways.Add(gateway);
                }

                LoggingService.Debug($"Carregados {Gateways.Count} gateways para coletor {coletorId}");
            }
            catch (Exception ex)
            {
                LoggingService.Error($"Erro ao carregar gateways para coletor {coletorId}", ex);
            }
            finally
            {
                IsLoadingFilters = false;
            }
        }

        /// <summary>
        /// Carrega equipamentos para um gateway
        /// </summary>
        private async Task LoadEquipmentsAsync(string coletorId, string gatewayId)
        {
            try
            {
                IsLoadingFilters = true;
                Equipments.Clear();
                Tags.Clear();

                var service = _configViewModel.GetInfluxDbService();
                if (service == null) return;

                var equipments = await service.GetAvailableEquipmentsAsync(coletorId, gatewayId);

                foreach (var equipment in equipments.OrderBy(e => e))
                {
                    Equipments.Add(equipment);
                }

                LoggingService.Debug($"Carregados {Equipments.Count} equipamentos");
            }
            catch (Exception ex)
            {
                LoggingService.Error($"Erro ao carregar equipamentos", ex);
            }
            finally
            {
                IsLoadingFilters = false;
            }
        }

        /// <summary>
        /// Carrega tags para um equipamento
        /// </summary>
        private async Task LoadTagsAsync(string coletorId, string gatewayId, string equipmentId)
        {
            try
            {
                IsLoadingFilters = true;
                Tags.Clear();

                var service = _configViewModel.GetInfluxDbService();
                if (service == null) return;

                var tags = await service.GetAvailableTagsAsync(coletorId, gatewayId, equipmentId);

                foreach (var tag in tags.OrderBy(t => t))
                {
                    Tags.Add(tag);
                }

                LoggingService.Debug($"Carregados {Tags.Count} tags");
            }
            catch (Exception ex)
            {
                LoggingService.Error($"Erro ao carregar tags", ex);
            }
            finally
            {
                IsLoadingFilters = false;
            }
        }

        #endregion

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
                var service = _configViewModel.GetInfluxDbService();

                if (service == null)
                {
                    StatusMessage = "Configure a conexão com o InfluxDB primeiro";
                    StatusMessageColor = Brushes.Orange;
                    return;
                }

                // Criar parâmetros de consulta
                var queryParams = new QueryParams
                {
                    ColetorId = SelectedColetor,
                    GatewayId = SelectedGateway,
                    EquipmentId = SelectedEquipment,
                    TagId = SelectedTag,
                    StartTime = StartDate,
                    EndTime = EndDate,
                    Limit = Limit
                };

                // Executar query
                var data = await service.QueryDataAsync(queryParams);

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

                StatusMessage = $"Consulta concluída: {Results.Count} registros encontrados";
                StatusMessageColor = Brushes.Green;
                LoggingService.Info($"Consulta retornou {Results.Count} registros");
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
                ExcelService.ExportToSheet(dataList, $"VortexData_{DateTime.Now:yyyyMMdd_HHmmss}");

                StatusMessage = $"Dados exportados para Excel: {Results.Count} registros";
                StatusMessageColor = Brushes.Green;
                LoggingService.Info($"Dados exportados para Excel: {Results.Count} registros");
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

        /// <summary>
        /// Comando para atualizar coletores
        /// </summary>
        [RelayCommand]
        private async Task RefreshColetoresAsync()
        {
            await LoadColetoresAsync();
        }

        #endregion
    }
}
