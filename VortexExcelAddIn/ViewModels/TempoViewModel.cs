using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Media;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using VortexExcelAddIn.Application.Factories;
using VortexExcelAddIn.Domain.Interfaces;
using VortexExcelAddIn.Domain.Models;
using VortexExcelAddIn.Models;
using VortexExcelAddIn.Services;

namespace VortexExcelAddIn.ViewModels
{
    /// <summary>
    /// ViewModel para o diálogo de configuração de agregação temporal.
    /// Segue o padrão MVVM e princípios SOLID (SRP, DIP).
    ///
    /// Responsabilidades:
    /// - Gerenciar estado da UI (seleções, mensagens de status)
    /// - Validar seleções do usuário
    /// - Executar estratégia de agregação apropriada
    /// - Atualizar resultados no QueryViewModel
    /// </summary>
    public partial class TempoViewModel : ViewModelBase
    {
        private readonly ConfigViewModel _configViewModel;
        private readonly QueryViewModel _queryViewModel;

        #region Observable Properties

        /// <summary>
        /// Lista de servidores disponíveis para seleção.
        /// </summary>
        [ObservableProperty]
        private ObservableCollection<ServerTypeItem> _availableServers;

        /// <summary>
        /// Servidor selecionado pelo usuário.
        /// </summary>
        [ObservableProperty]
        private ServerTypeItem _selectedServer;

        /// <summary>
        /// Tipo de servidor selecionado (contexto da UI).
        /// </summary>
        [ObservableProperty]
        private DatabaseType _serverType;

        /// <summary>
        /// Descrição do comportamento baseado no tipo de servidor.
        /// Historian: "Aplicar agregação aos dados brutos"
        /// VortexIO: "Filtrar dados já agregados"
        /// </summary>
        [ObservableProperty]
        private string _serverDescription;

        /// <summary>
        /// Lista de tipos de agregação disponíveis com checkbox.
        /// </summary>
        [ObservableProperty]
        private ObservableCollection<AggregationTypeItem> _availableAggregationTypes;

        /// <summary>
        /// Lista de janelas de tempo disponíveis com checkbox.
        /// </summary>
        [ObservableProperty]
        private ObservableCollection<TimeWindowItem> _availableTimeWindows;

        /// <summary>
        /// Mensagem de status exibida ao usuário.
        /// </summary>
        [ObservableProperty]
        private string _statusMessage;

        /// <summary>
        /// Cor da mensagem de status (Green = sucesso, Red = erro, Orange = aviso, Blue = processando).
        /// </summary>
        [ObservableProperty]
        private Brush _statusColor;

        /// <summary>
        /// Indica se uma operação está em andamento.
        /// </summary>
        [ObservableProperty]
        private bool _isProcessing;

        /// <summary>
        /// Token de autenticação do InfluxDB (editável no diálogo).
        /// </summary>
        [ObservableProperty]
        private string _token;

        /// <summary>
        /// Preview dos resultados da agregação (primeiros registros).
        /// </summary>
        [ObservableProperty]
        private ObservableCollection<VortexDataPoint> _previewResults;

        /// <summary>
        /// Resultados completos da agregação (não exibidos na aba Consultar Dados).
        /// </summary>
        private List<VortexDataPoint> _fullResults;

        /// <summary>
        /// Indica se há resultados disponíveis para exportar.
        /// </summary>
        [ObservableProperty]
        private bool _hasResults;

        /// <summary>
        /// Texto do botão de testar conexão.
        /// </summary>
        [ObservableProperty]
        private string _testConnectionButtonText;

        /// <summary>
        /// Cor de fundo do botão de testar.
        /// </summary>
        [ObservableProperty]
        private Brush _testConnectionButtonBackground;

        /// <summary>
        /// Indica se está testando conexão.
        /// </summary>
        [ObservableProperty]
        private bool _isTesting;

        #endregion

        /// <summary>
        /// Evento solicitando fechamento do diálogo.
        /// </summary>
        public event EventHandler RequestClose;

        public TempoViewModel(ConfigViewModel configViewModel, QueryViewModel queryViewModel)
        {
            _configViewModel = configViewModel ?? throw new ArgumentNullException(nameof(configViewModel));
            _queryViewModel = queryViewModel ?? throw new ArgumentNullException(nameof(queryViewModel));

            InitializeServers();
            InitializeAvailableOptions();
            LoadPreviousSelections();
            LoadTokenFromConfig();

            // Inicializar coleções
            PreviewResults = new ObservableCollection<VortexDataPoint>();
            _fullResults = new List<VortexDataPoint>();
            HasResults = false;

            // Inicializar botão de testar
            TestConnectionButtonText = "Testar Conexão";
            TestConnectionButtonBackground = new SolidColorBrush(Color.FromRgb(127, 127, 127)); // #7F7F7F - cinza

            StatusMessage = "Selecione o servidor, tipos de agregação e janelas de tempo";
            StatusColor = new SolidColorBrush(Color.FromRgb(149, 165, 166)); // #95A5A6 - cinza neutro

            LoggingService.Info("TempoViewModel inicializado");
        }

        /// <summary>
        /// Inicializa a lista de servidores disponíveis.
        /// </summary>
        private void InitializeServers()
        {
            AvailableServers = new ObservableCollection<ServerTypeItem>
            {
                new ServerTypeItem
                {
                    ServerType = DatabaseType.VortexHistorianAPI,
                    DisplayName = "Vortex Historian API",
                    Description = "🔄 Aplicar agregação em tempo real aos dados brutos usando Flux queries"
                },
                new ServerTypeItem
                {
                    ServerType = DatabaseType.VortexAPI,
                    DisplayName = "VortexIO API",
                    Description = "🔍 Filtrar dados já pré-agregados pelo Airflow (não re-agrega)"
                }
            };

            // Carregar último servidor selecionado ou usar o padrão do ConfigViewModel
            if (TempoConfiguration.LastSelectedServer.HasValue)
            {
                SelectedServer = AvailableServers.FirstOrDefault(
                    s => s.ServerType == TempoConfiguration.LastSelectedServer.Value);
            }
            else
            {
                // Usar o servidor atual do ConfigViewModel como padrão
                SelectedServer = AvailableServers.FirstOrDefault(
                    s => s.ServerType == _configViewModel.SelectedDatabaseType);
            }

            // Fallback para o primeiro servidor
            if (SelectedServer == null && AvailableServers.Count > 0)
            {
                SelectedServer = AvailableServers[0];
            }
        }

        /// <summary>
        /// Chamado quando o servidor selecionado muda.
        /// </summary>
        partial void OnSelectedServerChanged(ServerTypeItem value)
        {
            if (value != null)
            {
                ServerType = value.ServerType;
                ServerDescription = value.Description;
                LoggingService.Info($"[TempoViewModel] Servidor alterado para: {value.DisplayName}");
            }
        }

        /// <summary>
        /// Inicializa as opções disponíveis para seleção.
        /// </summary>
        private void InitializeAvailableOptions()
        {
            // Criar checkable items para tipos de agregação
            AvailableAggregationTypes = new ObservableCollection<AggregationTypeItem>
            {
                new AggregationTypeItem
                {
                    Type = VortexAggregationType.Average,
                    DisplayName = "Média (Average)",
                    IsSelected = false
                },
                new AggregationTypeItem
                {
                    Type = VortexAggregationType.Total,
                    DisplayName = "Total (Sum)",
                    IsSelected = false
                },
                new AggregationTypeItem
                {
                    Type = VortexAggregationType.MinMax,
                    DisplayName = "Mínimo/Máximo (Min/Max)",
                    IsSelected = false
                },
                new AggregationTypeItem
                {
                    Type = VortexAggregationType.FirstLast,
                    DisplayName = "Primeiro/Último (First/Last)",
                    IsSelected = false
                },
                new AggregationTypeItem
                {
                    Type = VortexAggregationType.Delta,
                    DisplayName = "Delta (Diferença)",
                    IsSelected = false
                }
            };

            // Criar checkable items para janelas de tempo
            AvailableTimeWindows = new ObservableCollection<TimeWindowItem>
            {
                new TimeWindowItem
                {
                    Window = TimeWindow.FiveMinutes,
                    DisplayName = "5 minutos",
                    IsSelected = false
                },
                new TimeWindowItem
                {
                    Window = TimeWindow.FifteenMinutes,
                    DisplayName = "15 minutos",
                    IsSelected = false
                },
                new TimeWindowItem
                {
                    Window = TimeWindow.ThirtyMinutes,
                    DisplayName = "30 minutos",
                    IsSelected = false
                },
                new TimeWindowItem
                {
                    Window = TimeWindow.SixtyMinutes,
                    DisplayName = "60 minutos (1 hora)",
                    IsSelected = false
                }
            };
        }

        /// <summary>
        /// Carrega as seleções anteriores do usuário.
        /// </summary>
        private void LoadPreviousSelections()
        {
            // Restaurar tipos de agregação selecionados
            foreach (var item in AvailableAggregationTypes)
            {
                if (TempoConfiguration.LastSelectedAggregationTypes.Contains(item.Type))
                {
                    item.IsSelected = true;
                }
            }

            // Restaurar janelas de tempo selecionadas
            foreach (var item in AvailableTimeWindows)
            {
                if (TempoConfiguration.LastSelectedTimeWindows.Contains(item.Window))
                {
                    item.IsSelected = true;
                }
            }

            LoggingService.Info($"[TempoViewModel] Seleções restauradas: {TempoConfiguration.LastSelectedAggregationTypes.Count} tipos, {TempoConfiguration.LastSelectedTimeWindows.Count} janelas");

            LoggingService.Info($"[TempoViewModel] Servidor: {ServerType}, Descrição: {ServerDescription}");
        }

        /// <summary>
        /// Carrega o token da configuração do ConfigViewModel.
        /// </summary>
        private void LoadTokenFromConfig()
        {
            try
            {
                // Tentar obter o token da configuração salva
                Token = _configViewModel.Token ?? string.Empty;
                LoggingService.Info($"[TempoViewModel] Token carregado da configuração: {(string.IsNullOrEmpty(Token) ? "vazio" : "preenchido")}");
            }
            catch (Exception ex)
            {
                LoggingService.Warn($"[TempoViewModel] Não foi possível carregar token da configuração: {ex.Message}");
                Token = string.Empty;
            }
        }

        /// <summary>
        /// Comando para testar a conexão com o InfluxDB.
        /// Faz um health check da API E uma query real para validar o token.
        /// </summary>
        [RelayCommand]
        private async Task TestConnection()
        {
            if (string.IsNullOrWhiteSpace(Token))
            {
                StatusMessage = "Informe o Token de autenticação para testar";
                StatusColor = new SolidColorBrush(Color.FromRgb(230, 126, 34)); // #E67E22 - laranja
                return;
            }

            IsTesting = true;
            TestConnectionButtonText = "Testando...";
            TestConnectionButtonBackground = new SolidColorBrush(Color.FromRgb(52, 152, 219)); // #3498DB - azul
            StatusMessage = "Testando conexão e validando token...";
            StatusColor = new SolidColorBrush(Color.FromRgb(52, 152, 219)); // #3498DB - azul

            try
            {
                var connection = CreateConnectionWithToken(Token);

                // 1. Test API connectivity (health check)
                var testResult = await connection.TestConnectionAsync();
                if (!testResult.IsSuccessful)
                {
                    StatusMessage = $"✗ API não acessível: {testResult.Message}";
                    StatusColor = new SolidColorBrush(Color.FromRgb(231, 76, 60)); // #E74C3C - vermelho
                    TestConnectionButtonText = "✗ API Offline";
                    TestConnectionButtonBackground = new SolidColorBrush(Color.FromRgb(231, 76, 60)); // #E74C3C - vermelho
                    LoggingService.Warn($"[TempoViewModel] API não acessível: {testResult.Message}");
                    return;
                }

                // 2. Test token by making a small query to validate InfluxDB credentials
                StatusMessage = "API online. Validando token no InfluxDB...";
                var testParams = new QueryParams
                {
                    StartTime = DateTime.UtcNow.AddMinutes(-5),
                    EndTime = DateTime.UtcNow,
                    Limit = 1 // Just need 1 record to validate token
                };

                try
                {
                    await connection.QueryDataAsync(testParams);

                    // If we get here, token is valid
                    var bucketName = ServerType == DatabaseType.VortexHistorianAPI ? "vortex_data" : "dados_airflow";
                    StatusMessage = $"✓ Token válido! Conexão com bucket '{bucketName}' confirmada.";
                    StatusColor = new SolidColorBrush(Color.FromRgb(39, 174, 96)); // #27AE60 - verde
                    TestConnectionButtonText = "✓ Token OK";
                    TestConnectionButtonBackground = new SolidColorBrush(Color.FromRgb(39, 174, 96)); // #27AE60 - verde
                    LoggingService.Info($"[TempoViewModel] Token validado com sucesso para bucket '{bucketName}'");
                }
                catch (Exception queryEx)
                {
                    // Token validation failed
                    var bucketName = ServerType == DatabaseType.VortexHistorianAPI ? "vortex_data" : "dados_airflow";
                    var errorMsg = queryEx.Message;

                    if (errorMsg.Contains("401") || errorMsg.Contains("unauthorized") || errorMsg.Contains("Unauthorized"))
                    {
                        StatusMessage = $"✗ Token inválido ou sem permissão para o bucket '{bucketName}'";
                    }
                    else
                    {
                        StatusMessage = $"✗ Erro ao validar token: {errorMsg}";
                    }

                    StatusColor = new SolidColorBrush(Color.FromRgb(231, 76, 60)); // #E74C3C - vermelho
                    TestConnectionButtonText = "✗ Token Inválido";
                    TestConnectionButtonBackground = new SolidColorBrush(Color.FromRgb(231, 76, 60)); // #E74C3C - vermelho
                    LoggingService.Warn($"[TempoViewModel] Token inválido para bucket '{bucketName}': {errorMsg}");
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Erro ao testar conexão: {ex.Message}";
                StatusColor = new SolidColorBrush(Color.FromRgb(231, 76, 60)); // #E74C3C - vermelho
                TestConnectionButtonText = "✗ Erro";
                TestConnectionButtonBackground = new SolidColorBrush(Color.FromRgb(231, 76, 60)); // #E74C3C - vermelho
                LoggingService.Error("[TempoViewModel] Erro ao testar conexão", ex);
            }
            finally
            {
                IsTesting = false;
                // Resetar botão após 3 segundos
                await Task.Delay(3000);
                TestConnectionButtonText = "Testar Conexão";
                TestConnectionButtonBackground = new SolidColorBrush(Color.FromRgb(127, 127, 127)); // #7F7F7F - cinza
            }
        }

        /// <summary>
        /// Comando para salvar a configuração (token) na configuração principal.
        /// </summary>
        [RelayCommand]
        private void SaveConfiguration()
        {
            if (string.IsNullOrWhiteSpace(Token))
            {
                StatusMessage = "Informe o Token antes de salvar";
                StatusColor = new SolidColorBrush(Color.FromRgb(230, 126, 34)); // #E67E22 - laranja
                return;
            }

            try
            {
                // Atualizar token no ConfigViewModel
                _configViewModel.Token = Token;

                StatusMessage = "✓ Configuração salva com sucesso!";
                StatusColor = new SolidColorBrush(Color.FromRgb(39, 174, 96)); // #27AE60 - verde
                LoggingService.Info("[TempoViewModel] Token salvo na configuração principal");
            }
            catch (Exception ex)
            {
                StatusMessage = $"Erro ao salvar configuração: {ex.Message}";
                StatusColor = new SolidColorBrush(Color.FromRgb(231, 76, 60)); // #E74C3C - vermelho
                LoggingService.Error("[TempoViewModel] Erro ao salvar configuração", ex);
            }
        }

        /// <summary>
        /// Comando para aplicar agregação/filtragem.
        /// </summary>
        [RelayCommand]
        private async Task ApplyAggregation()
        {
            // 1. Validar seleções
            var selectedTypes = AvailableAggregationTypes
                .Where(x => x.IsSelected)
                .Select(x => x.Type)
                .ToList();

            var selectedWindows = AvailableTimeWindows
                .Where(x => x.IsSelected)
                .Select(x => x.Window)
                .ToList();

            if (!selectedTypes.Any() || !selectedWindows.Any())
            {
                StatusMessage = "Selecione pelo menos um tipo de agregação e uma janela de tempo";
                StatusColor = new SolidColorBrush(Color.FromRgb(231, 76, 60)); // #E74C3C - vermelho
                LoggingService.Warn("[TempoViewModel] Validação falhou: seleções vazias");
                return;
            }

            // Validar token
            if (string.IsNullOrWhiteSpace(Token))
            {
                StatusMessage = "Informe o Token de autenticação do InfluxDB";
                StatusColor = new SolidColorBrush(Color.FromRgb(230, 126, 34)); // #E67E22 - laranja
                LoggingService.Warn("[TempoViewModel] Validação falhou: token vazio");
                return;
            }

            IsProcessing = true;
            StatusMessage = "Processando agregação...";
            StatusColor = new SolidColorBrush(Color.FromRgb(52, 152, 219)); // #3498DB - azul

            try
            {
                LoggingService.Info($"[TempoViewModel] Aplicando agregação: {selectedTypes.Count} tipos, {selectedWindows.Count} janelas");

                // 2. Criar configuração
                var config = new AggregationConfiguration
                {
                    AggregationTypes = selectedTypes,
                    TimeWindows = selectedWindows,
                    ServerType = ServerType
                };

                // 3. Criar conexão customizada com o token do diálogo
                IDataSourceConnection connection;
                try
                {
                    connection = CreateConnectionWithToken(Token);
                    LoggingService.Info($"[TempoViewModel] Conexão criada com token do diálogo");
                }
                catch (Exception ex)
                {
                    StatusMessage = $"Erro ao criar conexão: {ex.Message}";
                    StatusColor = new SolidColorBrush(Color.FromRgb(231, 76, 60)); // #E74C3C - vermelho
                    LoggingService.Error("[TempoViewModel] Erro ao criar conexão customizada", ex);
                    return;
                }

                // Verificar se o tipo de servidor suporta agregação
                if (!AggregationStrategyFactory.IsAggregationSupported(ServerType))
                {
                    StatusMessage = $"Agregação não suportada para {ServerType}";
                    StatusColor = new SolidColorBrush(Color.FromRgb(231, 76, 60)); // #E74C3C - vermelho
                    LoggingService.Warn($"[TempoViewModel] Agregação não suportada para {ServerType}");
                    return;
                }

                var strategy = AggregationStrategyFactory.CreateStrategy(ServerType, connection);
                LoggingService.Info($"[TempoViewModel] Estratégia criada: {strategy.GetType().Name}");

                // 4. Capturar parâmetros de query do QueryViewModel
                var baseParams = new QueryParams
                {
                    ColetorId = _queryViewModel.ColetorIds,
                    GatewayId = _queryViewModel.GatewayIds,
                    EquipmentId = _queryViewModel.EquipmentIds,
                    TagId = _queryViewModel.TagIds,
                    StartTime = _queryViewModel.StartDate,
                    EndTime = _queryViewModel.EndDate,
                    Limit = _queryViewModel.Limit
                };

                LoggingService.Info($"[TempoViewModel] Parâmetros: {baseParams.StartTime:yyyy-MM-dd} a {baseParams.EndTime:yyyy-MM-dd}, Limit={baseParams.Limit}");

                // 5. Executar agregação
                var results = await strategy.ApplyAggregationAsync(baseParams, config);

                LoggingService.Info($"[TempoViewModel] Agregação retornou {results.Count} pontos");

                // 6. Atualizar resultados LOCAIS (NÃO misturar com QueryViewModel)
                _fullResults = results;
                PreviewResults.Clear();

                // Preview (primeiros 20)
                foreach (var point in results.Take(20))
                {
                    PreviewResults.Add(point);
                }

                HasResults = results.Count > 0;

                StatusMessage = $"✓ Agregação concluída: {results.Count:N0} registros retornados";
                StatusColor = new SolidColorBrush(Color.FromRgb(39, 174, 96)); // #27AE60 - verde

                LoggingService.Info($"[TempoViewModel] Agregação aplicada com sucesso: {results.Count:N0} registros (mantidos separados da aba Consultar Dados)");

                // 7. Salvar seleções para próxima abertura
                SaveSelections(selectedTypes, selectedWindows);
            }
            catch (Exception ex)
            {
                StatusMessage = $"Erro ao aplicar agregação: {ex.Message}";
                StatusColor = new SolidColorBrush(Color.FromRgb(231, 76, 60)); // #E74C3C - vermelho
                LoggingService.Error("[TempoViewModel] Erro ao aplicar agregação", ex);
            }
            finally
            {
                IsProcessing = false;
            }
        }

        /// <summary>
        /// Cria uma conexão temporária usando o token fornecido.
        /// </summary>
        private IDataSourceConnection CreateConnectionWithToken(string token)
        {
            // Criar configuração temporária com o token do diálogo
            if (ServerType == DatabaseType.VortexHistorianAPI)
            {
                var config = new DataAccess.VortexAPI.HistorianApiConfig
                {
                    InfluxHost = "vortex_influxdb",
                    InfluxPort = 8086,
                    InfluxOrg = "vortex",
                    InfluxBucket = "vortex_data",
                    InfluxToken = token,
                    Timeout = 30
                };

                return new DataAccess.VortexAPI.HistorianApiDataSourceAdapter(config);
            }
            else if (ServerType == DatabaseType.VortexAPI)
            {
                var config = new DataAccess.VortexAPI.VortexApiConfig
                {
                    InfluxHost = "vortex_influxdb",
                    InfluxPort = 8086,
                    InfluxOrg = "vortex",
                    InfluxBucket = "dados_airflow",
                    InfluxToken = token,
                    Timeout = 30
                };

                return new DataAccess.VortexAPI.VortexApiDataSourceAdapter(config);
            }
            else
            {
                throw new NotSupportedException($"Tipo de servidor {ServerType} não suportado para agregação");
            }
        }

        /// <summary>
        /// Comando para exportar os resultados da agregação para Excel.
        /// </summary>
        [RelayCommand]
        private async Task ExportToExcel()
        {
            if (_fullResults == null || _fullResults.Count == 0)
            {
                StatusMessage = "Nenhum resultado para exportar. Execute a agregação primeiro.";
                StatusColor = new SolidColorBrush(Color.FromRgb(230, 126, 34)); // #E67E22 - laranja
                return;
            }

            IsProcessing = true;
            StatusMessage = "Exportando para Excel...";
            StatusColor = new SolidColorBrush(Color.FromRgb(52, 152, 219)); // #3498DB - azul

            try
            {
                await Task.Run(() =>
                {
                    ExcelService.ExportToSheet(_fullResults, null, ServerType);
                });

                StatusMessage = $"✓ {_fullResults.Count:N0} registros exportados para Excel com sucesso!";
                StatusColor = new SolidColorBrush(Color.FromRgb(39, 174, 96)); // #27AE60 - verde
                LoggingService.Info($"[TempoViewModel] {_fullResults.Count} registros da agregação exportados para Excel");

                // Fechar diálogo após exportação bem-sucedida
                await Task.Delay(1000);
                RequestClose?.Invoke(this, EventArgs.Empty);
            }
            catch (Exception ex)
            {
                StatusMessage = $"Erro ao exportar: {ex.Message}";
                StatusColor = new SolidColorBrush(Color.FromRgb(231, 76, 60)); // #E74C3C - vermelho
                LoggingService.Error("[TempoViewModel] Erro ao exportar para Excel", ex);
            }
            finally
            {
                IsProcessing = false;
            }
        }

        /// <summary>
        /// Salva as seleções atuais para persistir entre aberturas do diálogo.
        /// </summary>
        private void SaveSelections(
            List<VortexAggregationType> selectedTypes,
            List<TimeWindow> selectedWindows)
        {
            // Salvar servidor selecionado
            TempoConfiguration.LastSelectedServer = ServerType;

            // Salvar tipos de agregação
            TempoConfiguration.LastSelectedAggregationTypes.Clear();
            foreach (var type in selectedTypes)
            {
                TempoConfiguration.LastSelectedAggregationTypes.Add(type);
            }

            // Salvar janelas de tempo
            TempoConfiguration.LastSelectedTimeWindows.Clear();
            foreach (var window in selectedWindows)
            {
                TempoConfiguration.LastSelectedTimeWindows.Add(window);
            }

            LoggingService.Info($"[TempoViewModel] Seleções salvas: Servidor={ServerType}, {selectedTypes.Count} tipos, {selectedWindows.Count} janelas");
        }
    }

    /// <summary>
    /// Item de tipo de agregação para binding com checkbox.
    /// </summary>
    public partial class AggregationTypeItem : ObservableObject
    {
        [ObservableProperty]
        private VortexAggregationType _type;

        [ObservableProperty]
        private string _displayName;

        [ObservableProperty]
        private bool _isSelected;
    }

    /// <summary>
    /// Item de janela de tempo para binding com checkbox.
    /// </summary>
    public partial class TimeWindowItem : ObservableObject
    {
        [ObservableProperty]
        private TimeWindow _window;

        [ObservableProperty]
        private string _displayName;

        [ObservableProperty]
        private bool _isSelected;
    }

    /// <summary>
    /// Item de tipo de servidor para binding com ComboBox.
    /// </summary>
    public partial class ServerTypeItem : ObservableObject
    {
        [ObservableProperty]
        private DatabaseType _serverType;

        [ObservableProperty]
        private string _displayName;

        [ObservableProperty]
        private string _description;
    }
}
