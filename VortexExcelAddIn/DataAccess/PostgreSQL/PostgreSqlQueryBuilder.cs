using System;
using System.Linq;
using System.Text;
using VortexExcelAddIn.Domain.Interfaces;
using VortexExcelAddIn.Domain.Models;
using VortexExcelAddIn.Models;
using VortexExcelAddIn.Services;

namespace VortexExcelAddIn.DataAccess.PostgreSQL
{
    /// <summary>
    /// Constrói queries SQL para PostgreSQL.
    /// Implementa SRP (Single Responsibility Principle) - responsabilidade única de construir queries.
    /// </summary>
    public class PostgreSqlQueryBuilder : IQueryBuilder
    {
        private readonly PostgreSqlConfig _config;

        public PostgreSqlQueryBuilder(PostgreSqlConfig config)
        {
            _config = config ?? throw new ArgumentNullException(nameof(config));
        }

        /// <summary>
        /// Constrói query simples para teste de conexão.
        /// </summary>
        public string BuildTestQuery()
        {
            return "SELECT 1";
        }

        /// <summary>
        /// Constrói query para buscar dados com filtros.
        /// Usa parâmetros preparados para evitar SQL injection.
        /// </summary>
        public string BuildDataQuery(QueryParams parameters, TableSchema schema = null)
        {
            if (parameters == null)
                throw new ArgumentNullException(nameof(parameters));

            var tableSchema = schema ?? _config.TableSchema;
            var mapping = tableSchema.ColumnMapping;

            // Nome completo da tabela (schema.table)
            var fullTableName = string.IsNullOrEmpty(tableSchema.SchemaName)
                ? $"\"{tableSchema.TableName}\""
                : $"\"{tableSchema.SchemaName}\".\"{tableSchema.TableName}\"";

            var query = new StringBuilder();
            query.AppendLine($"SELECT");
            query.AppendLine($"    \"{mapping.TimeColumn}\" AS time,");
            query.AppendLine($"    \"{mapping.ValueColumn}\" AS valor,");
            query.AppendLine($"    \"{mapping.ColetorIdColumn}\" AS coletor_id,");
            query.AppendLine($"    \"{mapping.GatewayIdColumn}\" AS gateway_id,");
            query.AppendLine($"    \"{mapping.EquipmentIdColumn}\" AS equipment_id,");
            query.AppendLine($"    \"{mapping.TagIdColumn}\" AS tag_id");
            query.AppendLine($"FROM {fullTableName}");
            query.AppendLine($"WHERE \"{mapping.TimeColumn}\" BETWEEN @StartTime AND @EndTime");

            // Adicionar filtros opcionais (suporta múltiplos IDs separados por vírgula)
            if (!string.IsNullOrEmpty(parameters.ColetorId))
            {
                var filter = BuildMultiValueFilter(mapping.ColetorIdColumn, parameters.ColetorId, "ColetorId");
                query.AppendLine($"  AND ({filter})");
            }

            if (!string.IsNullOrEmpty(parameters.GatewayId))
            {
                var filter = BuildMultiValueFilter(mapping.GatewayIdColumn, parameters.GatewayId, "GatewayId");
                query.AppendLine($"  AND ({filter})");
            }

            if (!string.IsNullOrEmpty(parameters.EquipmentId))
            {
                var filter = BuildMultiValueFilter(mapping.EquipmentIdColumn, parameters.EquipmentId, "EquipmentId");
                query.AppendLine($"  AND ({filter})");
            }

            if (!string.IsNullOrEmpty(parameters.TagId))
            {
                var filter = BuildMultiValueFilter(mapping.TagIdColumn, parameters.TagId, "TagId");
                query.AppendLine($"  AND ({filter})");
            }

            query.AppendLine($"ORDER BY \"{mapping.TimeColumn}\" DESC");

            if (parameters.Limit.HasValue)
            {
                query.AppendLine($"LIMIT {parameters.Limit.Value}");
            }

            LoggingService.Debug($"Query PostgreSQL construída com sucesso");
            return query.ToString();
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

            var tableSchema = _config.TableSchema;
            var mapping = tableSchema.ColumnMapping;

            // Nome completo da tabela
            var fullTableName = string.IsNullOrEmpty(tableSchema.SchemaName)
                ? $"\"{tableSchema.TableName}\""
                : $"\"{tableSchema.SchemaName}\".\"{tableSchema.TableName}\"";

            // Converter tipo de agregação para função SQL
            var aggregationFunc = aggregation.ToSqlFunction();

            // Converter window period de formato "1m" para interval PostgreSQL "1 minute"
            var intervalStr = ConvertWindowPeriodToInterval(windowPeriod);

            var query = new StringBuilder();
            query.AppendLine($"SELECT");
            query.AppendLine($"    time_bucket('{intervalStr}', \"{mapping.TimeColumn}\") AS time,");
            query.AppendLine($"    {aggregationFunc}(\"{mapping.ValueColumn}\") AS valor,");
            query.AppendLine($"    \"{mapping.ColetorIdColumn}\" AS coletor_id,");
            query.AppendLine($"    \"{mapping.GatewayIdColumn}\" AS gateway_id,");
            query.AppendLine($"    \"{mapping.EquipmentIdColumn}\" AS equipment_id,");
            query.AppendLine($"    \"{mapping.TagIdColumn}\" AS tag_id");
            query.AppendLine($"FROM {fullTableName}");
            query.AppendLine($"WHERE \"{mapping.TimeColumn}\" BETWEEN @StartTime AND @EndTime");

            // Adicionar filtros opcionais
            if (!string.IsNullOrEmpty(parameters.ColetorId))
            {
                var filter = BuildMultiValueFilter(mapping.ColetorIdColumn, parameters.ColetorId, "ColetorId");
                query.AppendLine($"  AND ({filter})");
            }

            if (!string.IsNullOrEmpty(parameters.GatewayId))
            {
                var filter = BuildMultiValueFilter(mapping.GatewayIdColumn, parameters.GatewayId, "GatewayId");
                query.AppendLine($"  AND ({filter})");
            }

            if (!string.IsNullOrEmpty(parameters.EquipmentId))
            {
                var filter = BuildMultiValueFilter(mapping.EquipmentIdColumn, parameters.EquipmentId, "EquipmentId");
                query.AppendLine($"  AND ({filter})");
            }

            if (!string.IsNullOrEmpty(parameters.TagId))
            {
                var filter = BuildMultiValueFilter(mapping.TagIdColumn, parameters.TagId, "TagId");
                query.AppendLine($"  AND ({filter})");
            }

            query.AppendLine($"GROUP BY time, \"{mapping.ColetorIdColumn}\", \"{mapping.GatewayIdColumn}\", \"{mapping.EquipmentIdColumn}\", \"{mapping.TagIdColumn}\"");
            query.AppendLine($"ORDER BY time DESC");

            LoggingService.Debug($"Query agregada PostgreSQL construída: {aggregationFunc}, window: {intervalStr}");
            return query.ToString();
        }

        /// <summary>
        /// Valida se os parâmetros são válidos.
        /// </summary>
        public bool ValidateParameters(QueryParams parameters)
        {
            if (parameters == null)
                return false;

            return parameters.StartTime < parameters.EndTime;
        }

        /// <summary>
        /// Constrói filtro SQL para múltiplos valores separados por vírgula.
        /// Usa parâmetros preparados para evitar SQL injection.
        /// </summary>
        private string BuildMultiValueFilter(string columnName, string values, string parameterPrefix)
        {
            if (string.IsNullOrWhiteSpace(values))
                return "1=1";

            // Separar valores por vírgula e remover espaços em branco
            var valueList = values.Split(',')
                .Select(v => v.Trim())
                .Where(v => !string.IsNullOrEmpty(v))
                .ToList();

            if (valueList.Count == 0)
                return "1=1";

            if (valueList.Count == 1)
                return $"\"{columnName}\" = @{parameterPrefix}0";

            // Para múltiplos valores, usar IN
            var paramNames = string.Join(", ", valueList.Select((_, idx) => $"@{parameterPrefix}{idx}"));
            return $"\"{columnName}\" IN ({paramNames})";
        }

        /// <summary>
        /// Converte formato de período (ex: "1m", "5m", "1h") para interval PostgreSQL.
        /// </summary>
        private string ConvertWindowPeriodToInterval(string windowPeriod)
        {
            if (string.IsNullOrEmpty(windowPeriod))
                return "1 minute";

            // Extrair número e unidade (ex: "5m" -> 5, "m")
            var number = new string(windowPeriod.TakeWhile(char.IsDigit).ToArray());
            var unit = windowPeriod.Substring(number.Length).ToLower();

            var unitName = unit switch
            {
                "s" => "second",
                "m" => "minute",
                "h" => "hour",
                "d" => "day",
                _ => "minute"
            };

            return $"{number} {unitName}{(number != "1" ? "s" : "")}";
        }
    }
}
