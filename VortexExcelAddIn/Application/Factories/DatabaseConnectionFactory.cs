using System;
using System.Collections.Generic;
using VortexExcelAddIn.Application.Security;
using VortexExcelAddIn.DataAccess.InfluxDB;
using VortexExcelAddIn.DataAccess.PostgreSQL;
using VortexExcelAddIn.DataAccess.VortexAPI;
using VortexExcelAddIn.Domain.Interfaces;
using VortexExcelAddIn.Domain.Models;
using VortexExcelAddIn.Models;
using VortexExcelAddIn.Services;

namespace VortexExcelAddIn.Application.Factories
{
    /// <summary>
    /// Factory para criar conexões de banco de dados.
    /// Implementa OCP (Open/Closed Principle) - adicionar novos bancos sem modificar código existente.
    /// </summary>
    public class DatabaseConnectionFactory : IDatabaseConnectionFactory
    {
        private readonly ICredentialEncryptor _encryptor;
        private readonly Dictionary<DatabaseType, Func<UnifiedDatabaseConfig, IDataSourceConnection>> _factories;

        private DatabaseConnectionFactory(ICredentialEncryptor encryptor)
        {
            _encryptor = encryptor ?? throw new ArgumentNullException(nameof(encryptor));

            // Registrar factories para cada banco (OCP - extensível sem modificação)
            _factories = new Dictionary<DatabaseType, Func<UnifiedDatabaseConfig, IDataSourceConnection>>
            {
                { DatabaseType.VortexAPI, CreateVortexApiConnection },
                { DatabaseType.InfluxDB, CreateInfluxDbConnection },
                { DatabaseType.PostgreSQL, CreatePostgreSqlConnection }
                // MySQL, Oracle e SQL Server serão adicionados nas próximas fases
                // { DatabaseType.MySQL, CreateMySqlConnection },
                // { DatabaseType.Oracle, CreateOracleConnection },
                // { DatabaseType.SqlServer, CreateSqlServerConnection }
            };

            LoggingService.Info("DatabaseConnectionFactory inicializada");
        }

        /// <summary>
        /// Construtor padrão que usa DPAPICredentialEncryptor.
        /// </summary>
        public DatabaseConnectionFactory() : this(new DPAPICredentialEncryptor())
        {
        }

        /// <summary>
        /// Cria uma conexão baseada na configuração fornecida.
        /// </summary>
        public IDataSourceConnection CreateConnection(UnifiedDatabaseConfig config)
        {
            if (config == null)
                throw new ArgumentNullException(nameof(config));

            if (!config.IsValid())
                throw new ArgumentException("Configuração inválida. Verifique se todos os campos obrigatórios estão preenchidos.");

            if (!_factories.TryGetValue(config.DatabaseType, out var factory))
            {
                var supportedTypes = string.Join(", ", _factories.Keys);
                throw new NotSupportedException(
                    $"Banco de dados '{config.DatabaseType}' não está implementado ainda. " +
                    $"Tipos suportados: {supportedTypes}");
            }

            LoggingService.Info($"Criando conexão para banco: {config.DatabaseType}");
            return factory(config);
        }

        /// <summary>
        /// Cria configuração padrão para um tipo de banco.
        /// </summary>
        public UnifiedDatabaseConfig CreateDefaultConfig(DatabaseType type)
        {
            switch (type)
            {
                case DatabaseType.VortexAPI:
                    LoggingService.Debug("Criando configuração padrão para Vortex API (usa mesmas credenciais do InfluxDB)");
                    return new UnifiedDatabaseConfig
                    {
                        DatabaseType = DatabaseType.VortexAPI,
                        ConnectionSettings = new DatabaseConnectionSettings
                        {
                            Url = "http://localhost:8086",
                            Org = "vortex",
                            Bucket = "vortex_data",
                            EncryptedToken = string.Empty
                        }
                    };

                case DatabaseType.InfluxDB:
                    LoggingService.Debug("Criando configuração padrão para InfluxDB");
                    return new UnifiedDatabaseConfig
                    {
                        DatabaseType = DatabaseType.InfluxDB,
                        ConnectionSettings = new DatabaseConnectionSettings
                        {
                            Url = "http://localhost:8086",
                            Org = "vortex",
                            Bucket = "vortex_data",
                            EncryptedToken = string.Empty
                        }
                    };

                case DatabaseType.PostgreSQL:
                    LoggingService.Debug("Criando configuração padrão para PostgreSQL");
                    return new UnifiedDatabaseConfig
                    {
                        DatabaseType = DatabaseType.PostgreSQL,
                        ConnectionSettings = new DatabaseConnectionSettings
                        {
                            Host = "localhost",
                            Port = 5432,
                            DatabaseName = "vortex",
                            Username = "postgres",
                            EncryptedPassword = string.Empty,
                            UseSsl = false
                        },
                        TableSchema = new TableSchema
                        {
                            SchemaName = "public",
                            TableName = "dados_airflow"
                        }
                    };

                case DatabaseType.MySQL:
                    LoggingService.Debug("Criando configuração padrão para MySQL");
                    return new UnifiedDatabaseConfig
                    {
                        DatabaseType = DatabaseType.MySQL,
                        ConnectionSettings = new DatabaseConnectionSettings
                        {
                            Host = "localhost",
                            Port = 3306,
                            DatabaseName = "vortex",
                            Username = "root",
                            EncryptedPassword = string.Empty,
                            UseSsl = false
                        },
                        TableSchema = new TableSchema
                        {
                            SchemaName = string.Empty,
                            TableName = "dados_airflow"
                        }
                    };

                case DatabaseType.Oracle:
                    LoggingService.Debug("Criando configuração padrão para Oracle");
                    return new UnifiedDatabaseConfig
                    {
                        DatabaseType = DatabaseType.Oracle,
                        ConnectionSettings = new DatabaseConnectionSettings
                        {
                            Host = "localhost",
                            Port = 1521,
                            DatabaseName = "ORCL",
                            Username = "system",
                            EncryptedPassword = string.Empty,
                            UseSsl = false
                        },
                        TableSchema = new TableSchema
                        {
                            SchemaName = string.Empty,
                            TableName = "dados_airflow"
                        }
                    };

                case DatabaseType.SqlServer:
                    LoggingService.Debug("Criando configuração padrão para SQL Server");
                    return new UnifiedDatabaseConfig
                    {
                        DatabaseType = DatabaseType.SqlServer,
                        ConnectionSettings = new DatabaseConnectionSettings
                        {
                            Host = "localhost",
                            Port = 1433,
                            DatabaseName = "vortex",
                            Username = "sa",
                            EncryptedPassword = string.Empty,
                            UseSsl = false
                        },
                        TableSchema = new TableSchema
                        {
                            SchemaName = "dbo",
                            TableName = "dados_airflow"
                        }
                    };

                default:
                    throw new NotSupportedException($"Tipo de banco '{type}' não suportado");
            }
        }

        /// <summary>
        /// Verifica se um tipo de banco é suportado.
        /// </summary>
        public bool IsSupported(DatabaseType type)
        {
            return _factories.ContainsKey(type);
        }

        /// <summary>
        /// Cria conexão InfluxDB.
        /// </summary>
        private IDataSourceConnection CreateInfluxDbConnection(UnifiedDatabaseConfig config)
        {
            LoggingService.Info("Criando conexão InfluxDB");

            // Descriptografar token
            var token = _encryptor.Decrypt(config.ConnectionSettings.EncryptedToken);

            var influxConfig = new DataAccess.InfluxDB.InfluxDBConfig
            {
                Url = config.ConnectionSettings.Url,
                Token = token,
                Org = config.ConnectionSettings.Org,
                Bucket = config.ConnectionSettings.Bucket
            };

            // Criar componentes (SRP - cada um com sua responsabilidade)
            var queryBuilder = new InfluxDBQueryBuilder(influxConfig);
            var responseParser = new InfluxDBResponseParser();

            // Criar conexão com dependências injetadas (DIP)
            return new InfluxDBConnection(influxConfig, queryBuilder, responseParser);
        }

        /// <summary>
        /// Cria conexão PostgreSQL.
        /// </summary>
        private IDataSourceConnection CreatePostgreSqlConnection(UnifiedDatabaseConfig config)
        {
            LoggingService.Info("Criando conexão PostgreSQL");

            // Descriptografar senha
            var password = _encryptor.Decrypt(config.ConnectionSettings.EncryptedPassword);

            var pgConfig = new PostgreSqlConfig
            {
                Host = config.ConnectionSettings.Host,
                Port = config.ConnectionSettings.Port,
                Username = config.ConnectionSettings.Username,
                Password = password,
                DatabaseName = config.ConnectionSettings.DatabaseName,
                UseSsl = config.ConnectionSettings.UseSsl,
                TableSchema = config.TableSchema
            };

            // Criar componentes (SRP - cada um com sua responsabilidade)
            var queryBuilder = new PostgreSqlQueryBuilder(pgConfig);

            // Criar conexão com dependências injetadas (DIP)
            return new PostgreSqlConnection(pgConfig, queryBuilder);
        }

        /// <summary>
        /// Cria conexão via Vortex API com credenciais InfluxDB inline.
        /// </summary>
        private IDataSourceConnection CreateVortexApiConnection(UnifiedDatabaseConfig config)
        {
            LoggingService.Info("Criando conexão Vortex API com credenciais inline");

            // Extrair credenciais InfluxDB do config
            var url = config.ConnectionSettings.Url ?? "http://localhost:8086";
            var host = ExtractHostFromUrl(url);
            var port = ExtractPortFromUrl(url, 8086);

            // Descriptografar token
            var token = config.ConnectionSettings.EncryptedToken ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(token))
            {
                token = _encryptor.Decrypt(token);
            }

            var apiConfig = new VortexApiConfig
            {
                InfluxHost = host,
                InfluxPort = port,
                InfluxOrg = config.ConnectionSettings.Org ?? "vortex",
                InfluxBucket = config.ConnectionSettings.Bucket ?? "vortex_data",
                InfluxToken = token,
                Timeout = 30
            };

            // Criar adapter com credenciais inline
            return new VortexApiDataSourceAdapter(apiConfig);
        }

        /// <summary>
        /// Extrai o host de uma URL (ex: "http://localhost:8086" -> "localhost").
        /// </summary>
        private string ExtractHostFromUrl(string url)
        {
            if (string.IsNullOrWhiteSpace(url))
                return "localhost";

            try
            {
                var uri = new Uri(url);
                return uri.Host;
            }
            catch
            {
                // Se não for uma URL válida, assume que é só o host
                return url.Split(':')[0];
            }
        }

        /// <summary>
        /// Extrai a porta de uma URL (ex: "http://localhost:8086" -> 8086).
        /// </summary>
        private int ExtractPortFromUrl(string url, int defaultPort)
        {
            if (string.IsNullOrWhiteSpace(url))
                return defaultPort;

            try
            {
                var uri = new Uri(url);
                return uri.Port > 0 ? uri.Port : defaultPort;
            }
            catch
            {
                // Se não for uma URL válida, tenta extrair a porta manualmente
                var parts = url.Split(':');
                if (parts.Length > 1 && int.TryParse(parts[parts.Length - 1], out int port))
                {
                    return port;
                }
                return defaultPort;
            }
        }
    }
}
