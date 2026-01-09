using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using VortexExcelAddIn.Models;
using VortexExcelAddIn.Services;

namespace VortexExcelAddIn.DataAccess.VortexAPI
{
    /// <summary>
    /// Cliente HTTP para comunicação com a API Vortex IO.
    /// Implementa o padrão Singleton para reutilização do HttpClient.
    /// </summary>
    public class VortexApiClient : IDisposable
    {
        private readonly HttpClient _httpClient;
        private readonly string _baseUrl;
        private bool _disposed = false;

        /// <summary>
        /// Inicializa uma nova instância do VortexApiClient.
        /// </summary>
        /// <param name="baseUrl">URL base da API (ex: http://localhost:8000)</param>
        /// <param name="timeout">Timeout para requisições em segundos</param>
        public VortexApiClient(string baseUrl, int timeout = 30)
        {
            if (string.IsNullOrWhiteSpace(baseUrl))
                throw new ArgumentException("Base URL is required", nameof(baseUrl));

            _baseUrl = baseUrl.TrimEnd('/');
            _httpClient = new HttpClient
            {
                BaseAddress = new Uri(_baseUrl),
                Timeout = TimeSpan.FromSeconds(timeout)
            };

            _httpClient.DefaultRequestHeaders.Add("Accept", "application/json");

            LoggingService.Info($"VortexApiClient initialized with base URL: {_baseUrl}");
        }

        /// <summary>
        /// Testa a conexão com a API.
        /// </summary>
        /// <returns>True se a API estiver acessível, false caso contrário</returns>
        public async Task<(bool Success, string Message, TimeSpan Latency)> TestConnectionAsync()
        {
            var startTime = DateTime.UtcNow;

            try
            {
                var response = await _httpClient.GetAsync("/health");
                var latency = DateTime.UtcNow - startTime;

                if (response.IsSuccessStatusCode)
                {
                    LoggingService.Info($"API connection test successful - Latency: {latency.TotalMilliseconds:F0}ms");
                    return (true, "API connection successful", latency);
                }
                else
                {
                    var errorMessage = $"API returned status code: {response.StatusCode}";
                    LoggingService.Warn($"API connection test failed: {errorMessage}");
                    return (false, errorMessage, latency);
                }
            }
            catch (HttpRequestException ex)
            {
                var latency = DateTime.UtcNow - startTime;
                var errorMessage = $"HTTP request failed: {ex.Message}";
                LoggingService.Error($"API connection test failed: {errorMessage}", ex);
                return (false, errorMessage, latency);
            }
            catch (TaskCanceledException ex)
            {
                var latency = DateTime.UtcNow - startTime;
                var errorMessage = "Request timeout";
                LoggingService.Error($"API connection test failed: {errorMessage}", ex);
                return (false, errorMessage, latency);
            }
            catch (Exception ex)
            {
                var latency = DateTime.UtcNow - startTime;
                var errorMessage = $"Unexpected error: {ex.Message}";
                LoggingService.Error($"API connection test failed: {errorMessage}", ex);
                return (false, errorMessage, latency);
            }
        }

        /// <summary>
        /// Executa uma consulta de dados na API.
        /// </summary>
        /// <param name="request">Objeto de requisição com filtros de consulta</param>
        /// <returns>Lista de pontos de dados</returns>
        public async Task<List<VortexDataPoint>> QueryDataAsync(QueryRequestDto request)
        {
            try
            {
                var json = JsonConvert.SerializeObject(request);

                var content = new StringContent(json, Encoding.UTF8, "application/json");

                LoggingService.Debug($"Sending query request to API: {json}");

                var response = await _httpClient.PostAsync("/api/query", content);

                if (!response.IsSuccessStatusCode)
                {
                    var errorContent = await response.Content.ReadAsStringAsync();
                    throw new HttpRequestException(
                        $"API query failed with status {response.StatusCode}: {errorContent}");
                }

                var responseJson = await response.Content.ReadAsStringAsync();
                LoggingService.Debug($"Received query response from API (length: {responseJson.Length} chars)");

                var queryResponse = JsonConvert.DeserializeObject<QueryResponseDto>(responseJson);

                if (queryResponse == null || queryResponse.Data == null)
                {
                    LoggingService.Warn("API returned null or empty response");
                    return new List<VortexDataPoint>();
                }

                LoggingService.Info($"Query completed: {queryResponse.TotalCount} records returned in {queryResponse.QueryTimeMs:F0}ms");

                // Convert DTOs to domain models
                var dataPoints = queryResponse.Data.Select(dto => new VortexDataPoint(
                    time: dto.Time,
                    coletorId: dto.ColetorId,
                    gatewayId: dto.GatewayId,
                    equipmentId: dto.EquipmentId,
                    tagId: dto.TagId,
                    valor: dto.Valor
                )).ToList();

                return dataPoints;
            }
            catch (HttpRequestException ex)
            {
                LoggingService.Error($"HTTP error during query: {ex.Message}", ex);
                throw new InvalidOperationException($"Failed to query data from API: {ex.Message}", ex);
            }
            catch (JsonException ex)
            {
                LoggingService.Error($"JSON deserialization error: {ex.Message}", ex);
                throw new InvalidOperationException($"Failed to parse API response: {ex.Message}", ex);
            }
            catch (Exception ex)
            {
                LoggingService.Error($"Unexpected error during query: {ex.Message}", ex);
                throw;
            }
        }

        /// <summary>
        /// Obtém a lista de tags disponíveis para uma conexão.
        /// </summary>
        /// <param name="connectionId">ID da conexão no banco de dados da API</param>
        /// <returns>Lista de tags disponíveis</returns>
        public async Task<List<(string Id, string Name)>> GetTagsAsync(int connectionId)
        {
            try
            {
                var response = await _httpClient.GetAsync($"/api/tags/{connectionId}");

                if (!response.IsSuccessStatusCode)
                {
                    var errorContent = await response.Content.ReadAsStringAsync();
                    throw new HttpRequestException(
                        $"Failed to get tags with status {response.StatusCode}: {errorContent}");
                }

                var responseJson = await response.Content.ReadAsStringAsync();

                var tagsResponse = JsonConvert.DeserializeObject<TagsResponseDto>(responseJson);

                if (tagsResponse == null || tagsResponse.Tags == null)
                {
                    return new List<(string, string)>();
                }

                return tagsResponse.Tags.Select(t => (t.Id, t.Name)).ToList();
            }
            catch (Exception ex)
            {
                LoggingService.Error($"Error getting tags: {ex.Message}", ex);
                throw new InvalidOperationException($"Failed to get tags from API: {ex.Message}", ex);
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
                    LoggingService.Debug("VortexApiClient disposed");
                }

                _disposed = true;
            }
        }
    }

    #region DTOs

    /// <summary>
    /// DTO para credenciais InfluxDB inline (enviadas diretamente na requisição, sem ID de conexão).
    /// </summary>
    public class InfluxDBInlineCredentialsDto
    {
        [JsonProperty("host")]
        public string Host { get; set; }

        [JsonProperty("port")]
        public int Port { get; set; }

        [JsonProperty("org")]
        public string Org { get; set; }

        [JsonProperty("bucket")]
        public string Bucket { get; set; }

        [JsonProperty("token")]
        public string Token { get; set; }
    }

    /// <summary>
    /// DTO para requisição de consulta de dados.
    /// Suporta tanto connection_id (gerenciado) quanto inline_credentials (direto).
    /// </summary>
    public class QueryRequestDto
    {
        [JsonProperty("connection_id")]
        public int? ConnectionId { get; set; }

        [JsonProperty("inline_credentials")]
        public InfluxDBInlineCredentialsDto InlineCredentials { get; set; }

        [JsonProperty("measurement")]
        public string Measurement { get; set; }

        [JsonProperty("coletor_ids")]
        public List<string> ColetorIds { get; set; }

        [JsonProperty("gateway_ids")]
        public List<string> GatewayIds { get; set; }

        [JsonProperty("equipment_ids")]
        public List<string> EquipmentIds { get; set; }

        [JsonProperty("tag_ids")]
        public List<string> TagIds { get; set; }

        [JsonProperty("start_time")]
        public DateTime StartTime { get; set; }

        [JsonProperty("end_time")]
        public DateTime EndTime { get; set; }

        [JsonProperty("limit")]
        public int Limit { get; set; }
    }

    /// <summary>
    /// DTO para resposta de consulta de dados.
    /// </summary>
    public class QueryResponseDto
    {
        [JsonProperty("data")]
        public List<VortexDataPointDto> Data { get; set; }

        [JsonProperty("total_count")]
        public int TotalCount { get; set; }

        [JsonProperty("query_time_ms")]
        public double QueryTimeMs { get; set; }
    }

    /// <summary>
    /// DTO para ponto de dados individual.
    /// </summary>
    public class VortexDataPointDto
    {
        [JsonProperty("time")]
        public DateTime Time { get; set; }

        [JsonProperty("coletor_id")]
        public string ColetorId { get; set; }

        [JsonProperty("gateway_id")]
        public string GatewayId { get; set; }

        [JsonProperty("equipment_id")]
        public string EquipmentId { get; set; }

        [JsonProperty("tag_id")]
        public string TagId { get; set; }

        [JsonProperty("valor")]
        public string Valor { get; set; }
    }

    /// <summary>
    /// DTO para resposta de tags.
    /// </summary>
    public class TagsResponseDto
    {
        [JsonProperty("connection_id")]
        public int ConnectionId { get; set; }

        [JsonProperty("connection_name")]
        public string ConnectionName { get; set; }

        [JsonProperty("connection_type")]
        public string ConnectionType { get; set; }

        [JsonProperty("tags")]
        public List<TagDto> Tags { get; set; }

        [JsonProperty("count")]
        public int Count { get; set; }
    }

    /// <summary>
    /// DTO para tag individual.
    /// </summary>
    public class TagDto
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }
    }

    #endregion
}
