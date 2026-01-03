using System;
using System.Collections.ObjectModel;
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
    /// ViewModel para o painel de configuração.
    /// Refatorado para suportar múltiplos bancos de dados (SOLID: DIP).
    /// </summary>
    public partial class ConfigViewModel : ViewModelBase
    {
        private readonly IDatabaseConnectionFactory _connectionFactory;
        private IDataSourceConnection _dataSourceConnection;

        // Mantido para backward compatibility temporária
        private InfluxDbService _influxDbService;

        /// <summary>
        /// Evento disparado quando a conexão é salva e testada com sucesso
        /// </summary>
        public event EventHandler ConnectionSavedSuccessfully;

        #region Observable Properties

        // Campo para InfluxDB (apenas Token configurável)
        [ObservableProperty]
        private string _token;

        [ObservableProperty]
        private bool _isSaving;

        [ObservableProperty]
        private bool _isTesting;

        [ObservableProperty]
        private bool _isConnected;

        [ObservableProperty]
        private string _statusMessage;

        [ObservableProperty]
        private Brush _statusMessageColor;

        // Novas propriedades para suporte multi-banco
        [ObservableProperty]
        private DatabaseType _selectedDatabaseType;

        [ObservableProperty]
        private ObservableCollection<DatabaseTypeItem> _availableDatabaseTypes;

        [ObservableProperty]
        private bool _isRelationalDatabase;

        [ObservableProperty]
        private bool _isInfluxDbConnection;

        [ObservableProperty]
        private bool _isVortexApiConnection;

        // Campos para bancos relacionais
        [ObservableProperty]
        private string _host;

        [ObservableProperty]
        private int _port;

        [ObservableProperty]
        private string _username;

        [ObservableProperty]
        private string _password;

        [ObservableProperty]
        private string _databaseName;

        [ObservableProperty]
        private string _tableName;

        [ObservableProperty]
        private string _schemaName;

        [ObservableProperty]
        private bool _useSsl;

        // Propriedades para controlar o botão "Testar Conexão"
        [ObservableProperty]
        private string _testConnectionButtonText;

        [ObservableProperty]
        private Brush _testConnectionButtonBackground;

        [ObservableProperty]
        private Brush _testConnectionButtonForeground;

        #endregion

        public ConfigViewModel() : this(new DatabaseConnectionFactory())
        {
        }

        public ConfigViewModel(IDatabaseConnectionFactory connectionFactory)
        {
            _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));

            // Inicializar lista de tipos de banco disponíveis
            InitializeAvailableDatabaseTypes();

            // Carregar configuração do workbook ou usar padrão
            LoadConfiguration();

            // Inicializar cores de status
            StatusMessageColor = Brushes.Gray;
            UpdateStatusMessage();

            // Inicializar botão "Testar Conexão"
            TestConnectionButtonText = "Testar Conexão";
            TestConnectionButtonBackground = new SolidColorBrush(Color.FromRgb(127, 127, 127)); // #7F7F7F
            TestConnectionButtonForeground = Brushes.White;
        }

        /// <summary>
        /// Inicializa a lista de bancos de dados disponíveis.
        /// </summary>
        private void InitializeAvailableDatabaseTypes()
        {
            AvailableDatabaseTypes = new ObservableCollection<DatabaseTypeItem>
            {
                new DatabaseTypeItem { Type = DatabaseType.InfluxDB, DisplayName = "Servidor Vortex Historian" },
                new DatabaseTypeItem { Type = DatabaseType.VortexAPI, DisplayName = "Servidor VortexIO" }
            };
        }

        /// <summary>
        /// Carrega a configuração do workbook (v2 com migração automática).
        /// </summary>
        private void LoadConfiguration()
        {
            try
            {
                var config = ConfigService.LoadConfigV2();
                LoadFromUnifiedConfig(config);

                LoggingService.Debug($"Configuração v2 carregada no ViewModel (Tipo: {config.DatabaseType})");
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao carregar configuração no ViewModel", ex);
                SetDefaultConfig();
            }
        }

        /// <summary>
        /// Carrega propriedades do ViewModel a partir de UnifiedDatabaseConfig.
        /// </summary>
        private void LoadFromUnifiedConfig(UnifiedDatabaseConfig config)
        {
            if (config == null)
            {
                SetDefaultConfig();
                return;
            }

            SelectedDatabaseType = config.DatabaseType;

            if (config.DatabaseType == DatabaseType.VortexAPI)
            {
                Token = config.ConnectionSettings.EncryptedToken ?? string.Empty;
            }
            else if (config.DatabaseType == DatabaseType.InfluxDB)
            {
                Token = config.ConnectionSettings.EncryptedToken ?? string.Empty;
            }
            else
            {
                Host = config.ConnectionSettings.Host ?? string.Empty;
                Port = config.ConnectionSettings.Port;
                Username = config.ConnectionSettings.Username ?? string.Empty;
                Password = config.ConnectionSettings.EncryptedPassword ?? string.Empty;
                DatabaseName = config.ConnectionSettings.DatabaseName ?? string.Empty;
                UseSsl = config.ConnectionSettings.UseSsl;

                // Schema e tabela
                TableName = config.TableSchema?.TableName ?? string.Empty;
                SchemaName = config.TableSchema?.SchemaName ?? string.Empty;
            }

            UpdateFieldsVisibility();
        }

        /// <summary>
        /// Cria UnifiedDatabaseConfig a partir das propriedades do ViewModel.
        /// </summary>
        private UnifiedDatabaseConfig CreateUnifiedConfigFromViewModel()
        {
            var config = new UnifiedDatabaseConfig
            {
                DatabaseType = SelectedDatabaseType,
                ConfigVersion = 2
            };

            if (SelectedDatabaseType == DatabaseType.VortexAPI)
            {
                config.ConnectionSettings = new DatabaseConnectionSettings
                {
                    Url = "http://localhost:8086",
                    EncryptedToken = Token?.Trim() ?? string.Empty,
                    Org = "vortex",
                    Bucket = "dados_airflow"
                };
            }
            else if (SelectedDatabaseType == DatabaseType.InfluxDB)
            {
                config.ConnectionSettings = new DatabaseConnectionSettings
                {
                    Url = "http://localhost:8086",
                    EncryptedToken = Token?.Trim() ?? string.Empty,
                    Org = "vortex",
                    Bucket = "vortex_data"
                };
            }
            else
            {
                config.ConnectionSettings = new DatabaseConnectionSettings
                {
                    Host = Host?.Trim() ?? string.Empty,
                    Port = Port,
                    Username = Username?.Trim() ?? string.Empty,
                    EncryptedPassword = Password?.Trim() ?? string.Empty,
                    DatabaseName = DatabaseName?.Trim() ?? string.Empty,
                    UseSsl = UseSsl
                };

                config.TableSchema = new TableSchema
                {
                    TableName = TableName?.Trim() ?? string.Empty,
                    SchemaName = SchemaName?.Trim() ?? string.Empty
                };
            }

            return config;
        }

        /// <summary>
        /// Atualiza visibilidade dos campos baseado no tipo de banco selecionado.
        /// </summary>
        private void UpdateFieldsVisibility()
        {
            IsRelationalDatabase = SelectedDatabaseType.IsRelational();
            IsInfluxDbConnection = SelectedDatabaseType == DatabaseType.InfluxDB;
            IsVortexApiConnection = SelectedDatabaseType == DatabaseType.VortexAPI;
        }

        /// <summary>
        /// Atualiza mensagem de status baseado no tipo de banco selecionado.
        /// </summary>
        private void UpdateStatusMessage()
        {
            if (string.IsNullOrEmpty(StatusMessage) || StatusMessage.StartsWith("Configure"))
            {
                StatusMessage = $"Configure a Conexão com {SelectedDatabaseType.GetDisplayName()}";
            }
        }

        /// <summary>
        /// Chamado quando o tipo de banco é alterado.
        /// </summary>
        partial void OnSelectedDatabaseTypeChanged(DatabaseType value)
        {
            UpdateFieldsVisibility();
            UpdateStatusMessage();

            // Limpar conexão anterior
            _dataSourceConnection?.Dispose();
            _dataSourceConnection = null;

            // Definir valores padrão para o novo banco
            SetDefaultConfigForDatabaseType(value);

            LoggingService.Info($"Tipo de banco alterado para: {value}");
        }

        /// <summary>
        /// Define valores padrão para o banco atual.
        /// </summary>
        private void SetDefaultConfig()
        {
            SetDefaultConfigForDatabaseType(SelectedDatabaseType);
        }

        /// <summary>
        /// Define valores padrão para um tipo de banco específico.
        /// </summary>
        private void SetDefaultConfigForDatabaseType(DatabaseType databaseType)
        {
            var defaultConfig = _connectionFactory.CreateDefaultConfig(databaseType);
            LoadFromUnifiedConfig(defaultConfig);
        }

        /// <summary>
        /// Comando para salvar configuração (v2 - multi-banco).
        /// </summary>
        [RelayCommand]
        private async Task SaveAsync()
        {
            IsSaving = true;
            StatusMessage = "Salvando configuração...";
            StatusMessageColor = Brushes.Gray;

            try
            {
                // Criar config a partir do ViewModel
                var config = CreateUnifiedConfigFromViewModel();

                // Validar configuração
                if (!config.IsValid())
                {
                    StatusMessage = GetValidationErrorMessage(config);
                    StatusMessageColor = Brushes.Red;
                    return;
                }

                // Salvar no workbook (com criptografia automática)
                ConfigService.SaveConfigV2(config);

                // Testar conexão automaticamente após salvar
                await TestConnectionInternalAsync(config);

                if (IsConnected)
                {
                    StatusMessage = $"Configuração salva e conexão testada com sucesso! ({SelectedDatabaseType.GetDisplayName()})";
                    StatusMessageColor = Brushes.Green;
                    LoggingService.Info($"Configuração salva e testada: {SelectedDatabaseType}");

                    // Disparar evento para navegar automaticamente para a aba de Query
                    ConnectionSavedSuccessfully?.Invoke(this, EventArgs.Empty);
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Erro ao salvar configuração: {ex.Message}";
                StatusMessageColor = Brushes.Red;
                LoggingService.Error("Erro ao salvar configuração", ex);
                IsConnected = false;
            }
            finally
            {
                IsSaving = false;
            }
        }

        /// <summary>
        /// Obtém mensagem de erro de validação baseada no tipo de banco.
        /// </summary>
        private string GetValidationErrorMessage(UnifiedDatabaseConfig config)
        {
            if (config.DatabaseType == DatabaseType.VortexAPI)
            {
                if (string.IsNullOrWhiteSpace(config.ConnectionSettings.EncryptedToken))
                    return "Token é obrigatório";
                return null;
            }

            if (config.DatabaseType == DatabaseType.InfluxDB)
            {
                if (string.IsNullOrWhiteSpace(config.ConnectionSettings.EncryptedToken))
                    return "Token é obrigatório";
                return null;
            }

            if (string.IsNullOrWhiteSpace(config.ConnectionSettings.Host))
                return "Host é obrigatório";
            if (config.ConnectionSettings.Port <= 0)
                return "Porta é obrigatória";
            if (string.IsNullOrWhiteSpace(config.ConnectionSettings.DatabaseName))
                return "Nome do banco de dados é obrigatório";
            if (string.IsNullOrWhiteSpace(config.ConnectionSettings.Username))
                return "Usuário é obrigatório";

            return null;
        }

        /// <summary>
        /// Comando para testar conexão (v2 - multi-banco).
        /// </summary>
        [RelayCommand]
        private async Task TestConnectionAsync()
        {
            IsTesting = true;
            StatusMessage = $"Testando conexão com {SelectedDatabaseType.GetDisplayName()}...";
            StatusMessageColor = Brushes.Gray;

            try
            {
                var config = CreateUnifiedConfigFromViewModel();
                await TestConnectionInternalAsync(config);
            }
            catch (Exception ex)
            {
                StatusMessage = $"Erro ao testar conexão: {ex.Message}";
                StatusMessageColor = Brushes.Red;
                LoggingService.Error("Erro ao testar conexão", ex);
                IsConnected = false;
            }
            finally
            {
                IsTesting = false;
            }
        }

        /// <summary>
        /// Testa a conexão internamente usando a nova arquitetura.
        /// </summary>
        private async Task TestConnectionInternalAsync(UnifiedDatabaseConfig config)
        {
            try
            {
                // Criar conexão temporária para teste
                using (var testConnection = _connectionFactory.CreateConnection(config))
                {
                    var result = await testConnection.TestConnectionAsync();

                    if (result.IsSuccessful)
                    {
                        IsConnected = true;
                        StatusMessage = $"Conexão estabelecida com sucesso! ({SelectedDatabaseType.GetDisplayName()})";
                        if (result.Latency.TotalMilliseconds > 0)
                        {
                            StatusMessage += $"\nLatência: {result.Latency.TotalMilliseconds:F0}ms";
                        }
                        StatusMessageColor = Brushes.Green;

                        // Atualizar botão "Testar Conexão" para verde com "Conectado!"
                        TestConnectionButtonText = "Conectado!";
                        TestConnectionButtonBackground = Brushes.Green;
                        TestConnectionButtonForeground = Brushes.White;

                        // Atualizar conexão principal
                        _dataSourceConnection?.Dispose();
                        _dataSourceConnection = _connectionFactory.CreateConnection(config);

                        LoggingService.Info($"Conexão testada com sucesso: {SelectedDatabaseType}");
                    }
                    else
                    {
                        IsConnected = false;
                        StatusMessage = $"Falha na conexão: {result.Message}";
                        StatusMessageColor = Brushes.Red;

                        // Atualizar botão "Testar Conexão" para vermelho
                        TestConnectionButtonText = "Testar Conexão";
                        TestConnectionButtonBackground = Brushes.Red;
                        TestConnectionButtonForeground = Brushes.White;

                        LoggingService.Error($"Falha ao testar conexão: {result.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                IsConnected = false;
                StatusMessage = $"Erro: {ex.Message}";
                StatusMessageColor = Brushes.Red;

                // Atualizar botão "Testar Conexão" para vermelho
                TestConnectionButtonText = "Testar Conexão";
                TestConnectionButtonBackground = Brushes.Red;
                TestConnectionButtonForeground = Brushes.White;

                LoggingService.Error($"Falha ao testar conexão: {ex.Message}", ex);
                throw;
            }
        }

        /// <summary>
        /// Obtém o serviço InfluxDB configurado
        /// </summary>
        public InfluxDbService GetInfluxDbService()
        {
            if (_influxDbService == null)
            {
                // Verificar se há configuração válida
                if (string.IsNullOrWhiteSpace(Token))
                {
                    LoggingService.Warn("Tentativa de criar serviço InfluxDB sem configuração completa");
                    return null;
                }

                var config = new InfluxDBConfig
                {
                    Url = "http://localhost:8086",
                    Token = Token,
                    Org = "vortex",
                    Bucket = "vortex_data"
                };
                _influxDbService = new InfluxDbService(config);
                LoggingService.Info("Serviço InfluxDB criado com valores fixos");
            }

            return _influxDbService;
        }

        /// <summary>
        /// Obtém a conexão de banco de dados configurada (nova arquitetura SOLID).
        /// </summary>
        public IDataSourceConnection GetConnection()
        {
            if (_dataSourceConnection == null)
            {
                try
                {
                    var config = CreateUnifiedConfigFromViewModel();

                    if (!config.IsValid())
                    {
                        LoggingService.Warn("Tentativa de criar conexão sem configuração completa");
                        return null;
                    }

                    _dataSourceConnection = _connectionFactory.CreateConnection(config);
                    LoggingService.Info($"Conexão criada automaticamente: {config.DatabaseType}");
                }
                catch (Exception ex)
                {
                    LoggingService.Error("Erro ao criar conexão de banco de dados", ex);
                    return null;
                }
            }

            return _dataSourceConnection;
        }

        /// <summary>
        /// Cleanup
        /// </summary>
        public void Dispose()
        {
            _dataSourceConnection?.Dispose();
            _influxDbService?.Dispose();
        }
    }

    /// <summary>
    /// Item da lista de tipos de banco de dados para binding no ComboBox.
    /// </summary>
    public class DatabaseTypeItem
    {
        public DatabaseType Type { get; set; }
        public string DisplayName { get; set; }

        public override string ToString()
        {
            return DisplayName;
        }
    }
}
