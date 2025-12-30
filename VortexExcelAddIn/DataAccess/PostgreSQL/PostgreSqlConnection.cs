using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Npgsql;
using VortexExcelAddIn.Domain.Interfaces;
using VortexExcelAddIn.Domain.Models;
using VortexExcelAddIn.Models;
using VortexExcelAddIn.Services;

namespace VortexExcelAddIn.DataAccess.PostgreSQL
{
    /// <summary>
    /// Implementação de conexão com PostgreSQL.
    /// Implementa IDataSourceConnection (DIP) e ISupportsRawTableAccess (ISP).
    /// </summary>
    public class PostgreSqlConnection : IDataSourceConnection, ISupportsRawTableAccess
    {
        private readonly PostgreSqlConfig _config;
        private readonly PostgreSqlQueryBuilder _queryBuilder;
        private bool _disposed;

        public DatabaseType DatabaseType => DatabaseType.PostgreSQL;

        public PostgreSqlConnection(PostgreSqlConfig config, PostgreSqlQueryBuilder queryBuilder)
        {
            _config = config ?? throw new ArgumentNullException(nameof(config));
            _queryBuilder = queryBuilder ?? throw new ArgumentNullException(nameof(queryBuilder));

            LoggingService.Info($"PostgreSqlConnection inicializada: {_config.Host}:{_config.Port}/{_config.DatabaseName}");
        }

        /// <summary>
        /// Testa a conexão com o PostgreSQL.
        /// </summary>
        public async Task<ConnectionResult> TestConnectionAsync()
        {
            var startTime = DateTime.UtcNow;

            try
            {
                using (var connection = CreateConnection())
                {
                    await connection.OpenAsync();

                    using (var cmd = new NpgsqlCommand("SELECT version()", connection))
                    {
                        var version = await cmd.ExecuteScalarAsync();
                        var latency = DateTime.UtcNow - startTime;

                        LoggingService.Info($"Conexão com PostgreSQL testada com sucesso (latência: {latency.TotalMilliseconds}ms)");

                        return new ConnectionResult
                        {
                            IsSuccessful = true,
                            Message = "Conexão com PostgreSQL estabelecida com sucesso",
                            Latency = latency,
                            Metadata = new Dictionary<string, object>
                            {
                                { "Host", _config.Host },
                                { "Port", _config.Port },
                                { "Database", _config.DatabaseName },
                                { "Version", version?.ToString() ?? "Unknown" }
                            }
                        };
                    }
                }
            }
            catch (NpgsqlException pgEx)
            {
                LoggingService.Error($"Erro PostgreSQL ao conectar: {pgEx.Message}", pgEx);
                return ConnectionResult.Failure(
                    $"Falha na conexão PostgreSQL: {pgEx.Message}. Verifique host, porta, credenciais e se o PostgreSQL está rodando.",
                    pgEx);
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao testar conexão com PostgreSQL", ex);
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

            if (!_queryBuilder.ValidateParameters(parameters))
                throw new ArgumentException("Parâmetros inválidos: StartTime deve ser anterior a EndTime");

            try
            {
                var query = _queryBuilder.BuildDataQuery(parameters);
                LoggingService.Info($"Executando query PostgreSQL: {query}");

                using (var connection = CreateConnection())
                {
                    await connection.OpenAsync();

                    using (var cmd = new NpgsqlCommand(query, connection))
                    {
                        // Adicionar parâmetros preparados (proteção contra SQL injection)
                        AddQueryParameters(cmd, parameters);

                        var dataPoints = new List<VortexDataPoint>();

                        using (var reader = await cmd.ExecuteReaderAsync())
                        {
                            while (await reader.ReadAsync())
                            {
                                dataPoints.Add(MapDataPoint(reader));
                            }
                        }

                        LoggingService.Info($"Query retornou {dataPoints.Count} registros");
                        return dataPoints;
                    }
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao consultar dados do PostgreSQL", ex);
                throw new Exception($"Falha ao consultar dados: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Retorna informações sobre a conexão.
        /// </summary>
        public ConnectionInfo GetConnectionInfo()
        {
            return new ConnectionInfo
            {
                DatabaseType = DatabaseType.PostgreSQL,
                Host = $"{_config.Host}:{_config.Port}",
                DatabaseName = _config.DatabaseName,
                IsSecure = _config.UseSsl
            };
        }

        /// <summary>
        /// Retorna lista de schemas disponíveis no banco de dados.
        /// Implementa interface ISupportsRawTableAccess.
        /// </summary>
        public async Task<List<string>> GetAvailableSchemasAsync()
        {
            try
            {
                // language=sql
                var query = @"
                    SELECT schema_name
                    FROM information_schema.schemata
                    WHERE schema_name NOT IN ('pg_catalog', 'information_schema', 'pg_toast')
                    ORDER BY schema_name";

                using (var connection = CreateConnection())
                {
                    await connection.OpenAsync();

                    using (var cmd = new NpgsqlCommand(query, connection))
                    {
                        var schemas = new List<string>();

                        using (var reader = await cmd.ExecuteReaderAsync())
                        {
                            while (await reader.ReadAsync())
                            {
                                schemas.Add(reader.GetString(0));
                            }
                        }

                        LoggingService.Debug($"Encontrados {schemas.Count} schemas no PostgreSQL");
                        return schemas;
                    }
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao buscar schemas disponíveis", ex);
                throw;
            }
        }

        /// <summary>
        /// Retorna lista de tabelas em um schema específico.
        /// Implementa interface ISupportsRawTableAccess.
        /// </summary>
        public async Task<List<string>> GetTablesInSchemaAsync(string schemaName)
        {
            try
            {
                // language=sql
                var query = @"
                    SELECT table_name
                    FROM information_schema.tables
                    WHERE table_schema = @SchemaName
                      AND table_type = 'BASE TABLE'
                    ORDER BY table_name";

                using (var connection = CreateConnection())
                {
                    await connection.OpenAsync();

                    using (var cmd = new NpgsqlCommand(query, connection))
                    {
                        cmd.Parameters.AddWithValue("@SchemaName", schemaName ?? "public");

                        var tables = new List<string>();

                        using (var reader = await cmd.ExecuteReaderAsync())
                        {
                            while (await reader.ReadAsync())
                            {
                                tables.Add(reader.GetString(0));
                            }
                        }

                        LoggingService.Debug($"Encontradas {tables.Count} tabelas no schema '{schemaName}'");
                        return tables;
                    }
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error($"Erro ao buscar tabelas no schema '{schemaName}'", ex);
                throw;
            }
        }

        /// <summary>
        /// Obtém a estrutura (schema) de uma tabela específica.
        /// Implementa interface ISupportsRawTableAccess.
        /// </summary>
        public async Task<TableSchema> GetTableSchemaAsync(string tableName, string schemaName = null)
        {
            try
            {
                // language=sql
                var query = @"
                    SELECT column_name, data_type, is_nullable
                    FROM information_schema.columns
                    WHERE table_schema = @SchemaName
                      AND table_name = @TableName
                    ORDER BY ordinal_position";

                using (var connection = CreateConnection())
                {
                    await connection.OpenAsync();

                    using (var cmd = new NpgsqlCommand(query, connection))
                    {
                        cmd.Parameters.AddWithValue("@SchemaName", schemaName ?? "public");
                        cmd.Parameters.AddWithValue("@TableName", tableName);

                        var tableSchema = new TableSchema
                        {
                            SchemaName = schemaName ?? "public",
                            TableName = tableName,
                            ColumnMapping = new ColumnMapping()
                        };

                        using (var reader = await cmd.ExecuteReaderAsync())
                        {
                            // Tentar identificar as colunas automaticamente
                            while (await reader.ReadAsync())
                            {
                                var columnName = reader.GetString(0).ToLower();

                                // Mapear colunas conhecidas
                                if (columnName.Contains("time") || columnName.Contains("data") || columnName.Contains("timestamp"))
                                {
                                    tableSchema.ColumnMapping.TimeColumn = reader.GetString(0);
                                }
                                else if (columnName.Contains("valor") || columnName.Contains("value"))
                                {
                                    tableSchema.ColumnMapping.ValueColumn = reader.GetString(0);
                                }
                                else if (columnName.Contains("coletor"))
                                {
                                    tableSchema.ColumnMapping.ColetorIdColumn = reader.GetString(0);
                                }
                                else if (columnName.Contains("gateway"))
                                {
                                    tableSchema.ColumnMapping.GatewayIdColumn = reader.GetString(0);
                                }
                                else if (columnName.Contains("equipment") || columnName.Contains("equipamento"))
                                {
                                    tableSchema.ColumnMapping.EquipmentIdColumn = reader.GetString(0);
                                }
                                else if (columnName.Contains("tag"))
                                {
                                    tableSchema.ColumnMapping.TagIdColumn = reader.GetString(0);
                                }
                            }
                        }

                        LoggingService.Debug($"Schema obtido para tabela '{schemaName}.{tableName}'");
                        return tableSchema;
                    }
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error($"Erro ao obter schema da tabela '{schemaName}.{tableName}'", ex);
                throw;
            }
        }

        /// <summary>
        /// Cria uma nova conexão NpgsqlConnection.
        /// </summary>
        private NpgsqlConnection CreateConnection()
        {
            var connectionString = _config.BuildConnectionString();
            return new NpgsqlConnection(connectionString);
        }

        /// <summary>
        /// Adiciona parâmetros preparados ao comando SQL.
        /// IMPORTANTE: Sempre usar parâmetros preparados para evitar SQL injection.
        /// </summary>
        private void AddQueryParameters(NpgsqlCommand cmd, QueryParams parameters)
        {
            cmd.Parameters.AddWithValue("@StartTime", parameters.StartTime);
            cmd.Parameters.AddWithValue("@EndTime", parameters.EndTime);

            // Adicionar parâmetros para múltiplos valores
            if (!string.IsNullOrEmpty(parameters.ColetorId))
            {
                AddMultiValueParameters(cmd, parameters.ColetorId, "ColetorId");
            }

            if (!string.IsNullOrEmpty(parameters.GatewayId))
            {
                AddMultiValueParameters(cmd, parameters.GatewayId, "GatewayId");
            }

            if (!string.IsNullOrEmpty(parameters.EquipmentId))
            {
                AddMultiValueParameters(cmd, parameters.EquipmentId, "EquipmentId");
            }

            if (!string.IsNullOrEmpty(parameters.TagId))
            {
                AddMultiValueParameters(cmd, parameters.TagId, "TagId");
            }
        }

        /// <summary>
        /// Adiciona parâmetros para múltiplos valores separados por vírgula.
        /// </summary>
        private void AddMultiValueParameters(NpgsqlCommand cmd, string values, string parameterPrefix)
        {
            var valueList = values.Split(',')
                .Select(v => v.Trim())
                .Where(v => !string.IsNullOrEmpty(v))
                .ToList();

            for (int i = 0; i < valueList.Count; i++)
            {
                cmd.Parameters.AddWithValue($"@{parameterPrefix}{i}", valueList[i]);
            }
        }

        /// <summary>
        /// Mapeia um registro do DataReader para VortexDataPoint.
        /// </summary>
        private VortexDataPoint MapDataPoint(NpgsqlDataReader reader)
        {
            return new VortexDataPoint
            {
                Time = reader.GetDateTime(reader.GetOrdinal("time")),
                Valor = reader.IsDBNull(reader.GetOrdinal("valor")) ? null : reader.GetValue(reader.GetOrdinal("valor")).ToString(),
                ColetorId = reader.IsDBNull(reader.GetOrdinal("coletor_id")) ? null : reader.GetString(reader.GetOrdinal("coletor_id")),
                GatewayId = reader.IsDBNull(reader.GetOrdinal("gateway_id")) ? null : reader.GetString(reader.GetOrdinal("gateway_id")),
                EquipmentId = reader.IsDBNull(reader.GetOrdinal("equipment_id")) ? null : reader.GetString(reader.GetOrdinal("equipment_id")),
                TagId = reader.IsDBNull(reader.GetOrdinal("tag_id")) ? null : reader.GetString(reader.GetOrdinal("tag_id"))
            };
        }

        public void Dispose()
        {
            if (_disposed)
                return;

            // NpgsqlConnection é criada e descartada a cada operação (pattern using)
            // Não há recursos para liberar aqui
            LoggingService.Debug("PostgreSqlConnection disposed");
            _disposed = true;
        }
    }
}
