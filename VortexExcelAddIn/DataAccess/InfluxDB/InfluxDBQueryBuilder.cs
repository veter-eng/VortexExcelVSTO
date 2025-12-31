using System;
using System.Linq;
using VortexExcelAddIn.Domain.Interfaces;
using VortexExcelAddIn.Domain.Models;
using VortexExcelAddIn.Models;
using VortexExcelAddIn.Services;

namespace VortexExcelAddIn.DataAccess.InfluxDB
{
    /// <summary>
    /// Constrói queries Flux para InfluxDB.
    /// Implementa SRP (Single Responsibility Principle) - responsabilidade única de construir queries.
    /// Extraído do InfluxDBService para separação de responsabilidades.
    /// </summary>
    public class InfluxDBQueryBuilder : IQueryBuilder
    {
        private readonly InfluxDBConfig _config;

        public InfluxDBQueryBuilder(InfluxDBConfig config)
        {
            _config = config ?? throw new ArgumentNullException(nameof(config));
        }

        /// <summary>
        /// Constrói query simples para teste de conexão.
        /// </summary>
        public string BuildTestQuery()
        {
            var query = $@"
                from(bucket: ""{_config.Bucket}"")
                  |> range(start: -1m)
                  |> limit(n: 1)
            ";

            LoggingService.Debug($"Query de teste criada para bucket: {_config.Bucket}");
            return query;
        }

        /// <summary>
        /// Constrói query para buscar dados com filtros.
        /// </summary>
        public string BuildDataQuery(QueryParams parameters, TableSchema schema = null)
        {
            if (parameters == null)
                throw new ArgumentNullException(nameof(parameters));

            // Construir a query Flux
            // Filtrar apenas measurement "dados_rabbitmq" para garantir estrutura consistente
            var fluxQuery = $@"
                from(bucket: ""{_config.Bucket}"")
                  |> range(start: {FormatTimestamp(parameters.StartTime)}, stop: {FormatTimestamp(parameters.EndTime)})
                  |> filter(fn: (r) => r[""_measurement""] == ""dados_rabbitmq"")
                  |> filter(fn: (r) =>
                      exists r.coletor_id and r.coletor_id =~ /^[0-9]+$/ and
                      exists r.gateway_id and r.gateway_id =~ /^[0-9]+$/ and
                      exists r.equipment_id and r.equipment_id =~ /^[0-9]+$/ and
                      exists r.tag_id and r.tag_id =~ /^[0-9]+$/
                  )";

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

            LoggingService.Debug($"Query Flux construída com sucesso");
            return fluxQuery;
        }

        /// <summary>
        /// Constrói query com agregação de dados.
        /// </summary>
        public string BuildAggregatedQuery(
            QueryParams parameters,
            AggregationType aggregation,
            string windowPeriod = "1m")
        {
            if (parameters == null)
                throw new ArgumentNullException(nameof(parameters));

            // Converter enum para função Flux
            var aggregationFunc = aggregation.ToFluxFunction();

            var fluxQuery = $@"
                from(bucket: ""{_config.Bucket}"")
                  |> range(start: {FormatTimestamp(parameters.StartTime)}, stop: {FormatTimestamp(parameters.EndTime)})
                  |> filter(fn: (r) => r[""_measurement""] == ""dados_rabbitmq"")
                  |> filter(fn: (r) =>
                      exists r.coletor_id and r.coletor_id =~ /^[0-9]+$/ and
                      exists r.gateway_id and r.gateway_id =~ /^[0-9]+$/ and
                      exists r.equipment_id and r.equipment_id =~ /^[0-9]+$/ and
                      exists r.tag_id and r.tag_id =~ /^[0-9]+$/
                  )
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

            LoggingService.Debug($"Query agregada Flux construída: {aggregationFunc}, window: {windowPeriod}");
            return fluxQuery;
        }

        /// <summary>
        /// Valida se os parâmetros são válidos.
        /// </summary>
        public bool ValidateParameters(QueryParams parameters)
        {
            if (parameters == null)
                return false;

            // Validar se StartTime é anterior a EndTime
            return parameters.StartTime < parameters.EndTime;
        }

        /// <summary>
        /// Constrói filtro Flux para múltiplos valores separados por vírgula.
        /// Extraído do InfluxDBService linhas 82-102.
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
        /// Converte timestamp para formato RFC3339 (formato esperado pelo InfluxDB).
        /// Extraído do InfluxDBService linhas 298-300.
        /// </summary>
        private string FormatTimestamp(DateTime dateTime)
        {
            return dateTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ");
        }
    }
}
