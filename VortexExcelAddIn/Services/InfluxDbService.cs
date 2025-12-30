using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using VortexExcelAddIn.Domain.Models;
using VortexExcelAddIn.Models;

namespace VortexExcelAddIn.Services
{
    /// <summary>
    /// Serviço para comunicação com o InfluxDB via REST API
    /// </summary>
    public class InfluxDbService : IDisposable
    {
        private readonly HttpClient _httpClient;
        private readonly InfluxDBConfig _config;
        private bool _disposed;

        // Propriedades públicas para debug
        public string LastQueryExecuted { get; private set; }
        public string LastRawResponse { get; private set; }

        public InfluxDbService(InfluxDBConfig config)
        {
            _config = config ?? throw new ArgumentNullException(nameof(config));

            try
            {
                _httpClient = new HttpClient();
                _httpClient.BaseAddress = new Uri(_config.Url);
                _httpClient.Timeout = TimeSpan.FromMinutes(2); // Aumentar timeout para 2 minutos

                // Configurar autenticação com Token (InfluxDB v2 requer API Token)
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
            catch (HttpRequestException httpEx)
            {
                LoggingService.Error($"Erro HTTP ao conectar ao InfluxDB: {httpEx.Message}", httpEx);
                throw new Exception($"Falha na conexão HTTP: {httpEx.Message}. Verifique se o InfluxDB está rodando em {_config.Url}", httpEx);
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao testar conexão com InfluxDB", ex);
                throw new Exception($"Erro ao testar conexão: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Constrói filtro Flux para múltiplos valores separados por vírgula
        /// </summary>
        private string BuildMultiValueFilter(string fieldName, string values)
        {
            if (string.IsNullOrWhiteSpace(values))
                return "true";

            // Separar valores por vírgula e remover espaços em branco
            var valueList = values.Split(',')
                .Select(v => v.Trim())
                .Where(v => !string.IsNullOrEmpty(v))
                .ToList();

            if (valueList.Count == 0)
                return "true";

            if (valueList.Count == 1)
                return $@"r[""{fieldName}""] == ""{valueList[0]}""";

            // Para múltiplos valores, usar OR
            var conditions = valueList.Select(v => $@"r[""{fieldName}""] == ""{v}""");
            return string.Join(" or ", conditions);
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
                // Filtrar apenas measurement "dados_rabbitmq" para garantir estrutura consistente
                var fluxQuery = $@"
                    from(bucket: ""{_config.Bucket}"")
                      |> range(start: {FormatTimestamp(parameters.StartTime)}, stop: {FormatTimestamp(parameters.EndTime)})
                      |> filter(fn: (r) => r[""_measurement""] == ""dados_rabbitmq"")";
                // Adicionar filtros opcionais (suporta múltiplos IDs separados por vírgula)
                if (!string.IsNullOrEmpty(parameters.ColetorId))
                {
                    var filter = BuildMultiValueFilter("coletor_id", parameters.ColetorId);
                    fluxQuery += $@"
                      |> filter(fn: (r) => {filter})";
                }

                if (!string.IsNullOrEmpty(parameters.GatewayId))
                {
                    var filter = BuildMultiValueFilter("gateway_id", parameters.GatewayId);
                    fluxQuery += $@"
                      |> filter(fn: (r) => {filter})";
                }

                if (!string.IsNullOrEmpty(parameters.EquipmentId))
                {
                    var filter = BuildMultiValueFilter("equipment_id", parameters.EquipmentId);
                    fluxQuery += $@"
                      |> filter(fn: (r) => {filter})";
                }

                if (!string.IsNullOrEmpty(parameters.TagId))
                {
                    var filter = BuildMultiValueFilter("tag_id", parameters.TagId);
                    fluxQuery += $@"
                      |> filter(fn: (r) => {filter})";
                }

                // Consolidar todas as séries em uma única table antes de ordenar
                // Isso garante que teremos apenas um conjunto de headers
                fluxQuery += @"
                      |> group()
                      |> sort(columns: [""_time""], desc: true)";

                if (parameters.Limit.HasValue)
                {
                    fluxQuery += $@"
                      |> limit(n: {parameters.Limit.Value})";
                }

                // Salvar query para debug
                LastQueryExecuted = fluxQuery;
                LoggingService.Info($"Executando query Flux: {fluxQuery}");

                var response = await ExecuteFluxQueryAsync(fluxQuery);
                LastRawResponse = response;

                LoggingService.Info($"Resposta recebida do InfluxDB (primeiros 500 chars): {(response.Length > 500 ? response.Substring(0, 500) : response)}");

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

                // Adicionar filtros opcionais (suporta múltiplos IDs separados por vírgula)
                if (!string.IsNullOrEmpty(parameters.ColetorId))
                {
                    var filter = BuildMultiValueFilter("coletor_id", parameters.ColetorId);
                    fluxQuery += $@"
                      |> filter(fn: (r) => {filter})";
                }

                if (!string.IsNullOrEmpty(parameters.GatewayId))
                {
                    var filter = BuildMultiValueFilter("gateway_id", parameters.GatewayId);
                    fluxQuery += $@"
                      |> filter(fn: (r) => {filter})";
                }

                if (!string.IsNullOrEmpty(parameters.EquipmentId))
                {
                    var filter = BuildMultiValueFilter("equipment_id", parameters.EquipmentId);
                    fluxQuery += $@"
                      |> filter(fn: (r) => {filter})";
                }

                if (!string.IsNullOrEmpty(parameters.TagId))
                {
                    var filter = BuildMultiValueFilter("tag_id", parameters.TagId);
                    fluxQuery += $@"
                      |> filter(fn: (r) => {filter})";
                }

                fluxQuery += $@"
                      |> map(fn: (r) => ({{ r with _value: float(v: r._value) }}))
                      |> aggregateWindow(every: {windowPeriod}, fn: {aggregationFunc}, createEmpty: false)
                      |> sort(columns: [""_time""], desc: true)";

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
                LoggingService.Info($"Total de linhas na resposta CSV: {lines.Length}");

                if (lines.Length < 2)
                {
                    LoggingService.Warn("Resposta CSV vazia ou sem dados suficientes");
                    return dataPoints;
                }

                // Primeira linha é sempre o header (com ou sem #datatype antes)
                string[] headers = null;
                int dataStartIndex = 1;

                // Verificar se tem #datatype na primeira linha
                if (lines[0].StartsWith("#datatype"))
                {
                    // Formato com #datatype: próxima linha é o header
                    if (lines.Length > 2)
                    {
                        headers = lines[1].Split(',');
                        dataStartIndex = 2;
                    }
                }
                else
                {
                    // Formato direto: primeira linha já é o header
                    headers = lines[0].Split(',');
                    dataStartIndex = 1;
                }

                if (headers == null || headers.Length == 0)
                {
                    LoggingService.Warn("Nenhum header encontrado na resposta CSV");
                    return dataPoints;
                }

                LoggingService.Info($"Total de headers: {headers.Length}");
                LoggingService.Info($"Headers encontrados: [{string.Join("] [", headers)}]");
                LoggingService.Info($"Dados começam na linha: {dataStartIndex}");

                // Indices das colunas
                var timeIndex = Array.IndexOf(headers, "_time");
                var valueIndex = Array.IndexOf(headers, "_value");
                var coletorIndex = Array.IndexOf(headers, "coletor_id");
                var gatewayIndex = Array.IndexOf(headers, "gateway_id");
                var equipmentIndex = Array.IndexOf(headers, "equipment_id");
                var tagIndex = Array.IndexOf(headers, "tag_id");

                LoggingService.Info($"Índices das colunas - Time:{timeIndex}, Value:{valueIndex}, Coletor:{coletorIndex}, Gateway:{gatewayIndex}, Equipment:{equipmentIndex}, Tag:{tagIndex}");

                // Log das primeiras linhas de dados para debug
                if (lines.Length > dataStartIndex)
                {
                    for (int debugIdx = dataStartIndex; debugIdx < Math.Min(dataStartIndex + 3, lines.Length); debugIdx++)
                    {
                        var debugValues = lines[debugIdx].Split(',');
                        LoggingService.Info($"Linha {debugIdx} tem {debugValues.Length} valores: [{string.Join("] [", debugValues)}]");
                    }
                }

                // Parse dos dados
                int dataLineCount = 0;
                int invalidIdCount = 0;

                for (int i = dataStartIndex; i < lines.Length; i++)
                {
                    var line = lines[i];

                    // Pular apenas linhas completamente vazias
                    if (string.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }

                    var values = line.Split(',');

                    // Detectar linhas de header repetidas (quando há múltiplas tables)
                    // Headers sempre começam com vírgula vazia e contêm nomes de colunas do InfluxDB
                    bool isHeaderLine = false;

                    // Verificar se parece ser um header:
                    // 1. Começa com vírgula (primeira coluna vazia)
                    // 2. Contém palavras-chave de colunas do InfluxDB
                    if (line.StartsWith(",") && values.Length > 5)
                    {
                        // Verificar se contém colunas típicas do InfluxDB
                        var lineUpper = line.ToLower();
                        if (lineUpper.Contains("_time") && lineUpper.Contains("_value") &&
                            (lineUpper.Contains("_field") || lineUpper.Contains("_measurement")))
                        {
                            isHeaderLine = true;
                        }
                    }

                    if (isHeaderLine)
                    {
                        // Nova table detectada - atualizar headers
                        headers = values;

                        // Recalcular índices das colunas
                        timeIndex = Array.IndexOf(headers, "_time");
                        valueIndex = Array.IndexOf(headers, "_value");
                        coletorIndex = Array.IndexOf(headers, "coletor_id");
                        gatewayIndex = Array.IndexOf(headers, "gateway_id");
                        equipmentIndex = Array.IndexOf(headers, "equipment_id");
                        tagIndex = Array.IndexOf(headers, "tag_id");

                        LoggingService.Info($"***** NOVA TABLE DETECTADA na linha {i} ({dataLineCount} registros até agora) *****");
                        LoggingService.Info($"Headers: [{string.Join("] [", headers)}]");
                        LoggingService.Info($"Novos índices - Time:{timeIndex}, Value:{valueIndex}, Coletor:{coletorIndex}, Gateway:{gatewayIndex}, Equipment:{equipmentIndex}, Tag:{tagIndex}");
                        continue;
                    }

                    // Processar TODAS as linhas de dados, sem filtrar por número de colunas
                    // Se a linha tem pelo menos 1 valor, processar
                    if (values.Length == 0)
                    {
                        continue;
                    }

                    dataLineCount++;

                    bool isNearProblemArea = (dataLineCount >= 3995 && dataLineCount <= 4005);

                    if (isNearProblemArea)
                    {
                        LoggingService.Info($">>> Linha CSV {i}, DataPoint #{dataLineCount}: {values.Length} valores");
                        LoggingService.Info($">>> Linha bruta: {line.Substring(0, Math.Min(200, line.Length))}");
                        LoggingService.Info($">>> Índices atuais - Time:{timeIndex}, Value:{valueIndex}, Coletor:{coletorIndex}, Gateway:{gatewayIndex}, Equipment:{equipmentIndex}, Tag:{tagIndex}");
                    }

                    var dataPoint = new VortexDataPoint
                    {
                        Time = timeIndex >= 0 && timeIndex < values.Length ? ParseTimestamp(CleanCsvValue(values[timeIndex])) : DateTime.UtcNow,
                        Valor = valueIndex >= 0 && valueIndex < values.Length ? CleanCsvValue(values[valueIndex]) : string.Empty,
                        ColetorId = coletorIndex >= 0 && coletorIndex < values.Length ? CleanCsvValue(values[coletorIndex]) : string.Empty,
                        GatewayId = gatewayIndex >= 0 && gatewayIndex < values.Length ? CleanCsvValue(values[gatewayIndex]) : string.Empty,
                        EquipmentId = equipmentIndex >= 0 && equipmentIndex < values.Length ? CleanCsvValue(values[equipmentIndex]) : string.Empty,
                        TagId = tagIndex >= 0 && tagIndex < values.Length ? CleanCsvValue(values[tagIndex]) : string.Empty
                    };

                    // Validar se os IDs são numéricos ou vazios (filtrar valores inválidos como "dados_rabbitmq")
                    bool isValidColetorId = string.IsNullOrWhiteSpace(dataPoint.ColetorId) ||
                                           int.TryParse(dataPoint.ColetorId, out _);
                    bool isValidGatewayId = string.IsNullOrWhiteSpace(dataPoint.GatewayId) ||
                                           int.TryParse(dataPoint.GatewayId, out _);
                    bool isValidEquipmentId = string.IsNullOrWhiteSpace(dataPoint.EquipmentId) ||
                                             int.TryParse(dataPoint.EquipmentId, out _);
                    bool isValidTagId = string.IsNullOrWhiteSpace(dataPoint.TagId) ||
                                       int.TryParse(dataPoint.TagId, out _);

                    // Se algum ID for inválido, ignorar o registro
                    if (!isValidColetorId || !isValidGatewayId || !isValidEquipmentId || !isValidTagId)
                    {
                        invalidIdCount++;
                        if (invalidIdCount <= 10)  // Apenas os primeiros 10 para não poluir
                        {
                            LoggingService.Debug($"Linha {i} ignorada - IDs não numéricos: Coletor=[{dataPoint.ColetorId}], Gateway=[{dataPoint.GatewayId}], Equipment=[{dataPoint.EquipmentId}], Tag=[{dataPoint.TagId}]");
                        }
                        continue;
                    }

                    dataPoints.Add(dataPoint);

                    // Log detalhado dos primeiros registros para debug
                    if (dataLineCount <= 5 || isNearProblemArea)
                    {
                        LoggingService.Info($"DataPoint {dataLineCount} (linha CSV {i}): Time={dataPoint.Time:HH:mm:ss}, Valor=[{dataPoint.Valor}], Coletor=[{dataPoint.ColetorId}], Gateway=[{dataPoint.GatewayId}], Equipment=[{dataPoint.EquipmentId}], Tag=[{dataPoint.TagId}]");
                    }
                }

                LoggingService.Info($"Total de registros válidos: {dataLineCount}");
                if (invalidIdCount > 0)
                {
                    LoggingService.Info($"Total de registros filtrados (IDs não numéricos): {invalidIdCount}");
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

        /// <summary>
        /// Remove aspas duplas do início e fim de um valor CSV
        /// </summary>
        private string CleanCsvValue(string value)
        {
            if (string.IsNullOrEmpty(value))
                return value;

            // Remove aspas duplas do início e fim
            value = value.Trim();
            if (value.StartsWith("\"") && value.EndsWith("\"") && value.Length > 1)
            {
                value = value.Substring(1, value.Length - 2);
            }
            else if (value.StartsWith("\""))
            {
                value = value.Substring(1);
            }

            return value;
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

    // NOTA: Enum AggregationType movido para Domain/Models/AggregationType.cs
    // para evitar duplicação e suportar a nova arquitetura multi-banco.
}
