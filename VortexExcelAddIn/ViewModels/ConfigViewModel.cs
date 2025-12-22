using System;
using System.Threading.Tasks;
using System.Windows.Media;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using VortexExcelAddIn.Models;
using VortexExcelAddIn.Services;

namespace VortexExcelAddIn.ViewModels
{
    /// <summary>
    /// ViewModel para o painel de configuração
    /// Port do ConfigPanel.tsx
    /// </summary>
    public partial class ConfigViewModel : ViewModelBase
    {
        private InfluxDBService _influxDbService;

        #region Observable Properties

        [ObservableProperty]
        private string _url;

        [ObservableProperty]
        private string _token;

        [ObservableProperty]
        private string _org;

        [ObservableProperty]
        private string _bucket;

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

        #endregion

        public ConfigViewModel()
        {
            // Carregar configuração do workbook ou usar padrão
            LoadConfiguration();

            // Inicializar cores de status
            StatusMessageColor = Brushes.Gray;
            StatusMessage = "Configure a conexão com o InfluxDB";
        }

        /// <summary>
        /// Carrega a configuração do workbook
        /// </summary>
        private void LoadConfiguration()
        {
            try
            {
                var config = ConfigService.LoadConfig();
                Url = config.Url;
                Token = config.Token;
                Org = config.Org;
                Bucket = config.Bucket;

                LoggingService.Debug("Configuração carregada no ViewModel");
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao carregar configuração no ViewModel", ex);
                SetDefaultConfig();
            }
        }

        /// <summary>
        /// Define valores padrão
        /// </summary>
        private void SetDefaultConfig()
        {
            var defaultConfig = ConfigService.GetDefaultConfig();
            Url = defaultConfig.Url;
            Token = defaultConfig.Token;
            Org = defaultConfig.Org;
            Bucket = defaultConfig.Bucket;
        }

        /// <summary>
        /// Comando para salvar configuração
        /// </summary>
        [RelayCommand]
        private async Task SaveAsync()
        {
            IsSaving = true;
            StatusMessage = "Salvando configuração...";
            StatusMessageColor = Brushes.Gray;

            try
            {
                // Validar campos obrigatórios
                if (string.IsNullOrWhiteSpace(Url))
                {
                    StatusMessage = "URL é obrigatória";
                    StatusMessageColor = Brushes.Red;
                    return;
                }

                if (string.IsNullOrWhiteSpace(Token))
                {
                    StatusMessage = "Token é obrigatório";
                    StatusMessageColor = Brushes.Red;
                    return;
                }

                if (string.IsNullOrWhiteSpace(Org))
                {
                    StatusMessage = "Organização é obrigatória";
                    StatusMessageColor = Brushes.Red;
                    return;
                }

                if (string.IsNullOrWhiteSpace(Bucket))
                {
                    StatusMessage = "Bucket é obrigatório";
                    StatusMessageColor = Brushes.Red;
                    return;
                }

                // Criar config object
                var config = new InfluxDBConfig
                {
                    Url = Url.Trim(),
                    Token = Token.Trim(),
                    Org = Org.Trim(),
                    Bucket = Bucket.Trim()
                };

                // Salvar no workbook
                ConfigService.SaveConfig(config);

                // Testar conexão automaticamente após salvar
                await TestConnectionInternalAsync(config);

                if (IsConnected)
                {
                    StatusMessage = "Configuração salva e conexão testada com sucesso!";
                    StatusMessageColor = Brushes.Green;
                    LoggingService.Info("Configuração salva e testada com sucesso");
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
        /// Comando para testar conexão
        /// </summary>
        [RelayCommand]
        private async Task TestConnectionAsync()
        {
            IsTesting = true;
            StatusMessage = "Testando conexão...";
            StatusMessageColor = Brushes.Gray;

            try
            {
                var config = new InfluxDBConfig
                {
                    Url = Url?.Trim() ?? string.Empty,
                    Token = Token?.Trim() ?? string.Empty,
                    Org = Org?.Trim() ?? string.Empty,
                    Bucket = Bucket?.Trim() ?? string.Empty
                };

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
        /// Testa a conexão internamente
        /// </summary>
        private async Task TestConnectionInternalAsync(InfluxDBConfig config)
        {
            try
            {
                // Criar serviço temporário para teste
                using (var testService = new InfluxDBService(config))
                {
                    var result = await testService.TestConnectionAsync();

                    if (result)
                    {
                        IsConnected = true;
                        StatusMessage = "Conexão estabelecida com sucesso!";
                        StatusMessageColor = Brushes.Green;

                        // Atualizar serviço principal
                        _influxDbService?.Dispose();
                        _influxDbService = new InfluxDBService(config);
                    }
                    else
                    {
                        IsConnected = false;
                        StatusMessage = "Falha ao conectar. Verifique as credenciais.";
                        StatusMessageColor = Brushes.Red;
                    }
                }
            }
            catch (Exception ex)
            {
                IsConnected = false;
                StatusMessage = $"Erro de conexão: {ex.Message}";
                StatusMessageColor = Brushes.Red;
                throw;
            }
        }

        /// <summary>
        /// Obtém o serviço InfluxDB configurado
        /// </summary>
        public InfluxDBService GetInfluxDbService()
        {
            if (_influxDbService == null && IsConnected)
            {
                var config = new InfluxDBConfig
                {
                    Url = Url,
                    Token = Token,
                    Org = Org,
                    Bucket = Bucket
                };
                _influxDbService = new InfluxDBService(config);
            }

            return _influxDbService;
        }

        /// <summary>
        /// Cleanup
        /// </summary>
        public void Dispose()
        {
            _influxDbService?.Dispose();
        }
    }
}
