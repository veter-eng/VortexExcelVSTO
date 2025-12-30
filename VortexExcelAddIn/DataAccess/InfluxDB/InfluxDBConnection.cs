using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using VortexExcelAddIn.Domain.Interfaces;
using VortexExcelAddIn.Domain.Models;
using VortexExcelAddIn.Models;
using VortexExcelAddIn.Services;

namespace VortexExcelAddIn.DataAccess.InfluxDB
{
    /// <summary>
    /// Implementação de conexão com InfluxDB.
    /// Implementa IDataSourceConnection (DIP) e ISupportsAggregation (ISP).
    /// Substitui InfluxDBService com separação de responsabilidades (SRP).
    /// </summary>
    public class InfluxDBConnection : IDataSourceConnection, ISupportsAggregation
    {
        private readonly InfluxDBConfig _config;
        private readonly InfluxDBQueryBuilder _queryBuilder;
        private readonly InfluxDBResponseParser _responseParser;
        private readonly HttpClient _httpClient;
        private bool _disposed = false;

        // Propriedades públicas para debug
        public string LastQueryExecuted { get; private set; }
        public string LastRawResponse { get; private set; }

        public DatabaseType DatabaseType => DatabaseType.InfluxDB;

        public InfluxDBConnection(
            InfluxDBConfig config,
            InfluxDBQueryBuilder queryBuilder,
            InfluxDBResponseParser responseParser)
        {
            _config = config ?? throw new ArgumentNullException(nameof(config));
            _queryBuilder = queryBuilder ?? throw new ArgumentNullException(nameof(queryBuilder));
            _responseParser = responseParser ?? throw new ArgumentNullException(nameof(responseParser));

            try
            {
                _httpClient = new HttpClient
                {
                    BaseAddress = new Uri(_config.Url),
                    Timeout = TimeSpan.FromMinutes(2) // Timeout de 2 minutos
                };

                // Configurar autenticação com Token (InfluxDB v2 requer API Token)
                _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Token", _config.Token);
                _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                LoggingService.Info($"InfluxDBConnection inicializada: {_config.Url}");
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao inicializar InfluxDBConnection", ex);
                throw;
            }
        }

        /// <summary>
        /// Testa a conexão com o InfluxDB.
        /// </summary>
        public async Task<ConnectionResult> TestConnectionAsync()
        {
            var startTime = DateTime.UtcNow;

            try
            {
                var testQuery = _queryBuilder.BuildTestQuery();
                await ExecuteFluxQueryAsync(testQuery);

                var latency = DateTime.UtcNow - startTime;
                LoggingService.Info($"Conexão com InfluxDB testada com sucesso (latência: {latency.TotalMilliseconds}ms)");

                return new ConnectionResult
                {
                    IsSuccessful = true,
                    Message = "Conexão com InfluxDB estabelecida com sucesso",
                    Latency = latency,
                    Metadata = new Dictionary<string, object>
                    {
                        { "Url", _config.Url },
                        { "Org", _config.Org },
                        { "Bucket", _config.Bucket }
                    }
                };
            }
            catch (HttpRequestException httpEx)
            {
                LoggingService.Error($"Erro HTTP ao conectar ao InfluxDB: {httpEx.Message}", httpEx);
                return ConnectionResult.Failure(
                    $"Falha na conexão HTTP: {httpEx.Message}. Verifique se o InfluxDB está rodando em {_config.Url}",
                    httpEx);
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao testar conexão com InfluxDB", ex);
                return ConnectionResult.Failure($"Erro ao testar conexão: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Executa consulta e retorna dados.
        /// </summary>
        public async Task<List<VortexDataPoint>> QueryDataAsync(QueryParams parameters)
        {
            if (parameters == null)
                throw new ArgumentNullException(nameof(parameters));

            // Validar parâmetros
            if (!_queryBuilder.ValidateParameters(parameters))
                throw new ArgumentException("Parâmetros inválidos: StartTime deve ser anterior a EndTime");

            try
            {
                // Construir query usando o builder
                var fluxQuery = _queryBuilder.BuildDataQuery(parameters);

                // Salvar query para debug
                LastQueryExecuted = fluxQuery;
                LoggingService.Info($"Executando query Flux: {fluxQuery}");

                // Executar query
                var response = await ExecuteFluxQueryAsync(fluxQuery);
                LastRawResponse = response;

                LoggingService.Info($"Resposta recebida do InfluxDB (primeiros 500 chars): {(response.Length > 500 ? response.Substring(0, 500) : response)}");

                // Fazer parse da resposta usando o parser
                var dataPoints = _responseParser.Parse(response);

                LoggingService.Info($"Query retornou {dataPoints.Count} registros");
                return dataPoints;
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao consultar dados do InfluxDB", ex);
                throw new Exception($"Falha ao consultar dados: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Executa consulta com agregação de dados.
        /// Implementa interface ISupportsAggregation.
        /// </summary>
        public async Task<List<VortexDataPoint>> QueryAggregatedDataAsync(
            QueryParams parameters,
            AggregationType aggregation,
            string windowPeriod = "1m")
        {
            if (parameters == null)
                throw new ArgumentNullException(nameof(parameters));

            // Validar parâmetros
            if (!_queryBuilder.ValidateParameters(parameters))
                throw new ArgumentException("Parâmetros inválidos: StartTime deve ser anterior a EndTime");

            try
            {
                // Construir query agregada usando o builder
                var fluxQuery = _queryBuilder.BuildAggregatedQuery(parameters, aggregation, windowPeriod);

                // Salvar query para debug
                LastQueryExecuted = fluxQuery;
                LoggingService.Debug($"Executando query agregada: {fluxQuery}");

                // Executar query
                var response = await ExecuteFluxQueryAsync(fluxQuery);
                LastRawResponse = response;

                // Fazer parse da resposta usando o parser
                var dataPoints = _responseParser.Parse(response);

                LoggingService.Info($"Query agregada retornou {dataPoints.Count} registros");
                return dataPoints;
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao consultar dados agregados", ex);
                throw new Exception($"Falha ao consultar dados agregados: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Retorna informações sobre a conexão.
        /// </summary>
        public ConnectionInfo GetConnectionInfo()
        {
            return new ConnectionInfo
            {
                DatabaseType = DatabaseType.InfluxDB,
                Host = _config.Url,
                DatabaseName = _config.Bucket,
                IsSecure = _config.Url.StartsWith("https://", StringComparison.OrdinalIgnoreCase)
            };
        }

        /// <summary>
        /// Executa uma query Flux via REST API.
        /// </summary>
        private async Task<string> ExecuteFluxQueryAsync(string fluxQuery)
        {
            try
            {
                var requestBody = new
                {
                    query = fluxQuery,
                    type = "flux"
                };

                var json = JsonConvert.SerializeObject(requestBody);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                LoggingService.Debug($"Executando query para org: {_config.Org}");
                LoggingService.Debug($"URL Base: {_httpClient.BaseAddress}");
                LoggingService.Debug($"Auth Header: {_httpClient.DefaultRequestHeaders.Authorization?.Scheme} {(_httpClient.DefaultRequestHeaders.Authorization?.Parameter?.Length > 10 ? _httpClient.DefaultRequestHeaders.Authorization?.Parameter?.Substring(0, 10) + "..." : "null")}");

                var response = await _httpClient.PostAsync($"/api/v2/query?org={_config.Org}", content);

                if (!response.IsSuccessStatusCode)
                {
                    var errorContent = await response.Content.ReadAsStringAsync();
                    LoggingService.Error($"Erro HTTP {response.StatusCode}: {errorContent}");
                    throw new HttpRequestException($"HTTP {response.StatusCode}: {errorContent}");
                }

                return await response.Content.ReadAsStringAsync();
            }
            catch (Exception ex)
            {
                LoggingService.Error($"Erro ao executar query Flux: {ex.Message}", ex);
                throw;
            }
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
                    _httpClient?.Dispose();
                    LoggingService.Debug("InfluxDBConnection disposed");
                }

                _disposed = true;
            }
        }
    }
}
