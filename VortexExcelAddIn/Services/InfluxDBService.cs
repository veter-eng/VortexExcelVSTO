using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using VortexExcelAddIn.Models;

namespace VortexExcelAddIn.Services
{
    /// <summary>
    /// Serviço para comunicação com o InfluxDB via REST API
    /// </summary>
    public class InfluxDBService : IDisposable
    {
        private readonly HttpClient _httpClient;
        private readonly InfluxDBConfig _config;
        private bool _disposed = false;

        public InfluxDBService(InfluxDBConfig config)
        {
            _config = config ?? throw new ArgumentNullException(nameof(config));

            try
            {
                _httpClient = new HttpClient();
                _httpClient.BaseAddress = new Uri(_config.Url);
                _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Token", _config.Token);
                _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                LoggingService.Info($"InfluxDBService inicializado: {_config.Url}");
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao inicializar InfluxDBService", ex);
                throw;
            }
        }

        /// <summary>
        /// Testa a conexão com o InfluxDB
        /// </summary>
        public async Task<bool> TestConnectionAsync()
        {
            try
            {
                var fluxQuery = $@"
                    from(bucket: ""{_config.Bucket}"")
                      |> range(start: -1m)
                      |> limit(n: 1)
                ";

                await ExecuteFluxQueryAsync(fluxQuery);
                LoggingService.Info("Conexão com InfluxDB testada com sucesso");
                return true;
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao testar conexão com InfluxDB", ex);
                return false;
            }
        }

        /// <summary>
        /// Busca dados do InfluxDB com filtros
        /// </summary>
        public async Task<List<VortexDataPoint>> QueryDataAsync(QueryParams parameters)
        {
            if (parameters == null)
                throw new ArgumentNullException(nameof(parameters));

            var dataPoints = new List<VortexDataPoint>();

            try
            {
                // Construir a query Flux
                var fluxQuery = $@"
                    from(bucket: ""{_config.Bucket}"")
                      |> range(start: {FormatTimestamp(parameters.StartTime)}, stop: {FormatTimestamp(parameters.EndTime)})
                      |> filter(fn: (r) => r[""_measurement""] == ""dados_rabbitmq"")
                      |> filter(fn: (r) => r[""_field""] == ""valor"")";

                // Adicionar filtros opcionais
                if (!string.IsNullOrEmpty(parameters.ColetorId))
                {
                    fluxQuery += $@"
                      |> filter(fn: (r) => r[""coletor_id""] == ""{parameters.ColetorId}"")";
                }

                if (!string.IsNullOrEmpty(parameters.GatewayId))
                {
                    fluxQuery += $@"
                      |> filter(fn: (r) => r[""gateway_id""] == ""{parameters.GatewayId}"")";
                }

                if (!string.IsNullOrEmpty(parameters.EquipmentId))
                {
                    fluxQuery += $@"
                      |> filter(fn: (r) => r[""equipment_id""] == ""{parameters.EquipmentId}"")";
                }

                if (!string.IsNullOrEmpty(parameters.TagId))
                {
                    fluxQuery += $@"
                      |> filter(fn: (r) => r[""tag_id""] == ""{parameters.TagId}"")";
                }

                fluxQuery += @"
                      |> sort(columns: [""_time""])";

                if (parameters.Limit.HasValue)
                {
                    fluxQuery += $@"
                      |> limit(n: {parameters.Limit.Value})";
                }

                LoggingService.Debug($"Executando query Flux: {fluxQuery}");

                var response = await ExecuteFluxQueryAsync(fluxQuery);
                dataPoints = ParseFluxResponse(response);

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
        /// Busca lista de coletores disponíveis
        /// </summary>
        public async Task<List<string>> GetAvailableCollectorsAsync()
        {
            var fluxQuery = $@"
                from(bucket: ""{_config.Bucket}"")
                  |> range(start: -30d)
                  |> filter(fn: (r) => r[""_measurement""] == ""dados_rabbitmq"")
                  |> keep(columns: [""coletor_id""])
                  |> distinct(column: ""coletor_id"")
            ";

            try
            {
                var response = await ExecuteFluxQueryAsync(fluxQuery);
                var coletores = ParseDistinctValues(response, "coletor_id");

                LoggingService.Debug($"Encontrados {coletores.Count} coletores");
                return coletores;
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao buscar coletores", ex);
                return new List<string>();
            }
        }

        /// <summary>
        /// Busca lista de gateways disponíveis para um coletor
        /// </summary>
        public async Task<List<string>> GetAvailableGatewaysAsync(string coletorId)
        {
            if (string.IsNullOrEmpty(coletorId))
                return new List<string>();

            var fluxQuery = $@"
                from(bucket: ""{_config.Bucket}"")
                  |> range(start: -30d)
                  |> filter(fn: (r) => r[""_measurement""] == ""dados_rabbitmq"")
                  |> filter(fn: (r) => r[""coletor_id""] == ""{coletorId}"")
                  |> keep(columns: [""gateway_id""])
                  |> distinct(column: ""gateway_id"")
            ";

            try
            {
                var response = await ExecuteFluxQueryAsync(fluxQuery);
                var gateways = ParseDistinctValues(response, "gateway_id");

                LoggingService.Debug($"Encontrados {gateways.Count} gateways para coletor {coletorId}");
                return gateways;
            }
            catch (Exception ex)
            {
                LoggingService.Error($"Erro ao buscar gateways para coletor {coletorId}", ex);
                return new List<string>();
            }
        }

        /// <summary>
        /// Busca lista de equipamentos disponíveis para um gateway
        /// </summary>
        public async Task<List<string>> GetAvailableEquipmentsAsync(string coletorId, string gatewayId)
        {
            if (string.IsNullOrEmpty(coletorId) || string.IsNullOrEmpty(gatewayId))
                return new List<string>();

            var fluxQuery = $@"
                from(bucket: ""{_config.Bucket}"")
                  |> range(start: -30d)
                  |> filter(fn: (r) => r[""_measurement""] == ""dados_rabbitmq"")
                  |> filter(fn: (r) => r[""coletor_id""] == ""{coletorId}"")
                  |> filter(fn: (r) => r[""gateway_id""] == ""{gatewayId}"")
                  |> keep(columns: [""equipment_id""])
                  |> distinct(column: ""equipment_id"")
            ";

            try
            {
                var response = await ExecuteFluxQueryAsync(fluxQuery);
                var equipments = ParseDistinctValues(response, "equipment_id");

                LoggingService.Debug($"Encontrados {equipments.Count} equipamentos");
                return equipments;
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao buscar equipamentos", ex);
                return new List<string>();
            }
        }

        /// <summary>
        /// Busca lista de tags disponíveis para um equipamento
        /// </summary>
        public async Task<List<string>> GetAvailableTagsAsync(string coletorId, string gatewayId, string equipmentId)
        {
            if (string.IsNullOrEmpty(coletorId) || string.IsNullOrEmpty(gatewayId) || string.IsNullOrEmpty(equipmentId))
                return new List<string>();

            var fluxQuery = $@"
                from(bucket: ""{_config.Bucket}"")
                  |> range(start: -30d)
                  |> filter(fn: (r) => r[""_measurement""] == ""dados_rabbitmq"")
                  |> filter(fn: (r) => r[""coletor_id""] == ""{coletorId}"")
                  |> filter(fn: (r) => r[""gateway_id""] == ""{gatewayId}"")
                  |> filter(fn: (r) => r[""equipment_id""] == ""{equipmentId}"")
                  |> keep(columns: [""tag_id""])
                  |> distinct(column: ""tag_id"")
            ";

            try
            {
                var response = await ExecuteFluxQueryAsync(fluxQuery);
                var tags = ParseDistinctValues(response, "tag_id");

                LoggingService.Debug($"Encontrados {tags.Count} tags");
                return tags;
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao buscar tags", ex);
                return new List<string>();
            }
        }

        /// <summary>
        /// Busca dados agregados (média, min, max, count)
        /// </summary>
        public async Task<List<VortexDataPoint>> QueryAggregatedDataAsync(
            QueryParams parameters,
            AggregationType aggregation,
            string windowPeriod = "1m")
        {
            if (parameters == null)
                throw new ArgumentNullException(nameof(parameters));

            var dataPoints = new List<VortexDataPoint>();

            try
            {
                var aggregationFunc = aggregation.ToString().ToLower();

                var fluxQuery = $@"
                    from(bucket: ""{_config.Bucket}"")
                      |> range(start: {FormatTimestamp(parameters.StartTime)}, stop: {FormatTimestamp(parameters.EndTime)})
                      |> filter(fn: (r) => r[""_measurement""] == ""dados_rabbitmq"")
                      |> filter(fn: (r) => r[""_field""] == ""valor"")";

                if (!string.IsNullOrEmpty(parameters.ColetorId))
                {
                    fluxQuery += $@"
                      |> filter(fn: (r) => r[""coletor_id""] == ""{parameters.ColetorId}"")";
                }

                if (!string.IsNullOrEmpty(parameters.GatewayId))
                {
                    fluxQuery += $@"
                      |> filter(fn: (r) => r[""gateway_id""] == ""{parameters.GatewayId}"")";
                }

                if (!string.IsNullOrEmpty(parameters.EquipmentId))
                {
                    fluxQuery += $@"
                      |> filter(fn: (r) => r[""equipment_id""] == ""{parameters.EquipmentId}"")";
                }

                if (!string.IsNullOrEmpty(parameters.TagId))
                {
                    fluxQuery += $@"
                      |> filter(fn: (r) => r[""tag_id""] == ""{parameters.TagId}"")";
                }

                fluxQuery += $@"
                      |> map(fn: (r) => ({{ r with _value: float(v: r._value) }}))
                      |> aggregateWindow(every: {windowPeriod}, fn: {aggregationFunc}, createEmpty: false)
                      |> sort(columns: [""_time""])";

                LoggingService.Debug($"Executando query agregada: {fluxQuery}");

                var response = await ExecuteFluxQueryAsync(fluxQuery);
                dataPoints = ParseFluxResponse(response);

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
        /// Executa uma query Flux via REST API
        /// </summary>
        private async Task<string> ExecuteFluxQueryAsync(string fluxQuery)
        {
            var requestBody = new
            {
                query = fluxQuery,
                type = "flux",
                org = _config.Org
            };

            var json = JsonConvert.SerializeObject(requestBody);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await _httpClient.PostAsync("/api/v2/query", content);
            response.EnsureSuccessStatusCode();

            return await response.Content.ReadAsStringAsync();
        }

        /// <summary>
        /// Converte timestamp para formato RFC3339
        /// </summary>
        private string FormatTimestamp(DateTime dateTime)
        {
            return dateTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ");
        }

        /// <summary>
        /// Parse da resposta CSV do InfluxDB
        /// </summary>
        private List<VortexDataPoint> ParseFluxResponse(string csvResponse)
        {
            var dataPoints = new List<VortexDataPoint>();

            try
            {
                var lines = csvResponse.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);

                // Procurar linhas de dados (pulando cabeçalhos e anotações)
                var dataStartIndex = -1;
                string[] headers = null;

                for (int i = 0; i < lines.Length; i++)
                {
                    if (lines[i].StartsWith("#datatype"))
                    {
                        // Próxima linha após #datatype é o cabeçalho
                        if (i + 1 < lines.Length)
                        {
                            headers = lines[i + 1].Split(',');
                            dataStartIndex = i + 2;
                            break;
                        }
                    }
                }

                if (dataStartIndex == -1 || headers == null)
                    return dataPoints;

                // Indices das colunas
                var timeIndex = Array.IndexOf(headers, "_time");
                var valueIndex = Array.IndexOf(headers, "_value");
                var coletorIndex = Array.IndexOf(headers, "coletor_id");
                var gatewayIndex = Array.IndexOf(headers, "gateway_id");
                var equipmentIndex = Array.IndexOf(headers, "equipment_id");
                var tagIndex = Array.IndexOf(headers, "tag_id");

                // Parse dos dados
                for (int i = dataStartIndex; i < lines.Length; i++)
                {
                    var values = lines[i].Split(',');
                    if (values.Length < headers.Length)
                        continue;

                    var dataPoint = new VortexDataPoint
                    {
                        Time = timeIndex >= 0 && timeIndex < values.Length ? ParseTimestamp(values[timeIndex]) : DateTime.UtcNow,
                        Valor = valueIndex >= 0 && valueIndex < values.Length ? values[valueIndex] : string.Empty,
                        ColetorId = coletorIndex >= 0 && coletorIndex < values.Length ? values[coletorIndex] : string.Empty,
                        GatewayId = gatewayIndex >= 0 && gatewayIndex < values.Length ? values[gatewayIndex] : string.Empty,
                        EquipmentId = equipmentIndex >= 0 && equipmentIndex < values.Length ? values[equipmentIndex] : string.Empty,
                        TagId = tagIndex >= 0 && tagIndex < values.Length ? values[tagIndex] : string.Empty
                    };

                    dataPoints.Add(dataPoint);
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao fazer parse da resposta Flux", ex);
            }

            return dataPoints;
        }

        /// <summary>
        /// Parse de valores distintos da resposta CSV
        /// </summary>
        private List<string> ParseDistinctValues(string csvResponse, string columnName)
        {
            var values = new List<string>();

            try
            {
                var lines = csvResponse.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);

                var dataStartIndex = -1;
                string[] headers = null;

                for (int i = 0; i < lines.Length; i++)
                {
                    if (lines[i].StartsWith("#datatype"))
                    {
                        if (i + 1 < lines.Length)
                        {
                            headers = lines[i + 1].Split(',');
                            dataStartIndex = i + 2;
                            break;
                        }
                    }
                }

                if (dataStartIndex == -1 || headers == null)
                    return values;

                var columnIndex = Array.IndexOf(headers, columnName);
                if (columnIndex == -1)
                    return values;

                for (int i = dataStartIndex; i < lines.Length; i++)
                {
                    var cols = lines[i].Split(',');
                    if (columnIndex < cols.Length && !string.IsNullOrEmpty(cols[columnIndex]))
                    {
                        values.Add(cols[columnIndex]);
                    }
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error($"Erro ao fazer parse de valores distintos para {columnName}", ex);
            }

            return values.Distinct().ToList();
        }

        /// <summary>
        /// Parse de timestamp RFC3339
        /// </summary>
        private DateTime ParseTimestamp(string timestamp)
        {
            if (DateTime.TryParse(timestamp, out var result))
                return result.ToUniversalTime();

            return DateTime.UtcNow;
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
                    LoggingService.Debug("InfluxDBService disposed");
                }

                _disposed = true;
            }
        }
    }

    /// <summary>
    /// Tipos de agregação suportados
    /// </summary>
    public enum AggregationType
    {
        Mean,
        Min,
        Max,
        Count
    }
}
