using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using VortexExcelAddIn.Domain.Interfaces;
using VortexExcelAddIn.Domain.Models;
using VortexExcelAddIn.Models;
using VortexExcelAddIn.Services;

namespace VortexExcelAddIn.DataAccess.VortexAPI
{
    /// <summary>
    /// Adapter para integração da API Vortex IO com a interface IDataSourceConnection.
    /// Implementa o padrão Adapter (GoF Design Pattern).
    /// Segue o princípio DIP (Dependency Inversion Principle) do SOLID.
    /// Usa credenciais inline (enviadas diretamente na requisição) ao invés de ID de conexão gerenciado.
    /// </summary>
    public class VortexApiDataSourceAdapter : IDataSourceConnection
    {
        private readonly VortexApiConfig _config;
        private readonly VortexApiClient _apiClient;
        private bool _disposed = false;

        /// <summary>
        /// Tipo de banco de dados desta conexão (sempre VortexAPI).
        /// </summary>
        public DatabaseType DatabaseType => DatabaseType.VortexAPI;

        /// <summary>
        /// Inicializa uma nova instância do adapter.
        /// </summary>
        /// <param name="config">Configuração da API com credenciais InfluxDB inline</param>
        public VortexApiDataSourceAdapter(VortexApiConfig config)
        {
            _config = config ?? throw new ArgumentNullException(nameof(config));

            if (!_config.IsValid())
            {
                throw new ArgumentException("Invalid API configuration: missing InfluxDB credentials", nameof(config));
            }

            _apiClient = new VortexApiClient(_config.ApiUrl, _config.Timeout);

            LoggingService.Info($"VortexApiDataSourceAdapter initialized with inline credentials for {_config.InfluxHost}:{_config.InfluxPort}");
        }

        /// <summary>
        /// Testa a conectividade com a API.
        /// </summary>
        /// <returns>Resultado do teste de conexão</returns>
        public async Task<ConnectionResult> TestConnectionAsync()
        {
            try
            {
                var (success, message, latency) = await _apiClient.TestConnectionAsync();

                if (success)
                {
                    return new ConnectionResult
                    {
                        IsSuccessful = true,
                        Message = $"Conectado ao Servidor VortexIO - InfluxDB: {_config.InfluxHost}:{_config.InfluxPort}",
                        Latency = latency,
                        Metadata = new Dictionary<string, object>
                        {
                            { "api_url", _config.ApiUrl },
                            { "influx_host", _config.InfluxHost },
                            { "influx_port", _config.InfluxPort },
                            { "influx_org", _config.InfluxOrg },
                            { "influx_bucket", _config.InfluxBucket }
                        }
                    };
                }
                else
                {
                    return ConnectionResult.Failure(message);
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error($"API connection test failed: {ex.Message}", ex);
                return ConnectionResult.Failure($"Failed to connect to API: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Executa consulta de dados através da API usando credenciais inline.
        /// </summary>
        /// <param name="parameters">Parâmetros de consulta</param>
        /// <returns>Lista de pontos de dados</returns>
        public async Task<List<VortexDataPoint>> QueryDataAsync(QueryParams parameters)
        {
            if (parameters == null)
            {
                throw new ArgumentNullException(nameof(parameters));
            }

            try
            {
                // Mapear QueryParams para QueryRequestDto com credenciais inline
                var request = new QueryRequestDto
                {
                    // Não usa ConnectionId - envia credenciais inline
                    ConnectionId = null,
                    InlineCredentials = new InfluxDBInlineCredentialsDto
                    {
                        // Converte localhost para host.docker.internal pois a API roda em container Docker
                        Host = _config.InfluxHost == "localhost" ? "host.docker.internal" : _config.InfluxHost,
                        Port = _config.InfluxPort,
                        Org = _config.InfluxOrg,
                        Bucket = _config.InfluxBucket,
                        Token = _config.InfluxToken
                    },
                    Measurement = "dados_airflow", // VortexIO usa dados processados do Airflow
                    ColetorIds = ParseCsvToList(parameters.ColetorId),
                    GatewayIds = ParseCsvToList(parameters.GatewayId),
                    EquipmentIds = ParseCsvToList(parameters.EquipmentId),
                    TagIds = ParseCsvToList(parameters.TagId),
                    StartTime = parameters.StartTime,
                    EndTime = parameters.EndTime,
                    Limit = parameters.Limit ?? 1000
                };

                LoggingService.Info(
                    $"[VORTEXIO DEBUG] Querying VortexIO API with measurement=dados_airflow - Coletor: {parameters.ColetorId}, " +
                    $"Gateway: {parameters.GatewayId}, Equipment: {parameters.EquipmentId}, " +
                    $"Tag: {parameters.TagId}, Time range: {parameters.StartTime:yyyy-MM-dd HH:mm} to {parameters.EndTime:yyyy-MM-dd HH:mm}");

                var dataPoints = await _apiClient.QueryDataAsync(request);

                LoggingService.Info($"Query completed: {dataPoints.Count} records returned from VortexIO API");

                return dataPoints;
            }
            catch (InvalidOperationException ex)
            {
                // Re-throw API-specific errors
                LoggingService.Error($"VortexIO API query failed: {ex.Message}", ex);
                throw;
            }
            catch (Exception ex)
            {
                LoggingService.Error($"Unexpected error during VortexIO API query: {ex.Message}", ex);
                throw new InvalidOperationException($"Failed to query data from VortexIO API: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Retorna informações sobre a conexão atual.
        /// </summary>
        /// <returns>Informações da conexão</returns>
        public ConnectionInfo GetConnectionInfo()
        {
            return new ConnectionInfo
            {
                DatabaseType = DatabaseType.VortexAPI,
                Host = $"{_config.ApiUrl} -> {_config.InfluxHost}:{_config.InfluxPort}",
                DatabaseName = $"{_config.InfluxOrg}/{_config.InfluxBucket}",
                Username = "VortexIO API",
                IsSecure = _config.ApiUrl.StartsWith("https://", StringComparison.OrdinalIgnoreCase),
                ServerVersion = "Vortex IO API v2.0 (Inline Credentials)"
            };
        }

        /// <summary>
        /// Converte uma string CSV em lista de strings.
        /// Se a string for vazia ou null, retorna null (sem filtros).
        /// </summary>
        /// <param name="csv">String separada por vírgulas</param>
        /// <returns>Lista de strings ou null</returns>
        private List<string> ParseCsvToList(string csv)
        {
            if (string.IsNullOrWhiteSpace(csv))
            {
                return null; // Null significa "sem filtro" na API
            }

            return csv.Split(',')
                .Select(s => s.Trim())
                .Where(s => !string.IsNullOrEmpty(s))
                .ToList();
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    _apiClient?.Dispose();
                    LoggingService.Debug("VortexApiDataSourceAdapter disposed");
                }

                _disposed = true;
            }
        }
    }
}
