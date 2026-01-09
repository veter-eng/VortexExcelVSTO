using System.Xml.Serialization;
using VortexExcelAddIn.Domain.Models;

namespace VortexExcelAddIn.Models
{
    /// <summary>
    /// Configuração unificada que suporta todos os tipos de banco de dados.
    /// Substitui a antiga InfluxDBConfig com suporte a múltiplos bancos (SOLID: OCP).
    /// </summary>
    [XmlRoot("UnifiedDatabaseConfig", Namespace = "http://vortex.com/database-config-v2")]
    public class UnifiedDatabaseConfig
    {
        /// <summary>
        /// Tipo de banco de dados configurado.
        /// </summary>
        public DatabaseType DatabaseType { get; set; }

        /// <summary>
        /// Configurações de conexão (credenciais, host, porta, etc.).
        /// </summary>
        public DatabaseConnectionSettings ConnectionSettings { get; set; }

        /// <summary>
        /// Configuração de tabela/schema (apenas para bancos relacionais).
        /// Ignorado para InfluxDB.
        /// </summary>
        public TableSchema TableSchema { get; set; }

        /// <summary>
        /// Versão da configuração (para migrações futuras).
        /// </summary>
        public int ConfigVersion { get; set; }

        public UnifiedDatabaseConfig()
        {
            DatabaseType = DatabaseType.VortexHistorianAPI; // padrão
            ConnectionSettings = new DatabaseConnectionSettings();
            TableSchema = new TableSchema();
            ConfigVersion = 2;
        }

        /// <summary>
        /// Valida se a configuração está completa e consistente.
        /// </summary>
        /// <returns>True se válida, False caso contrário</returns>
        public bool IsValid()
        {
            if (ConnectionSettings == null)
                return false;

            // Validações específicas por tipo de banco
            switch (DatabaseType)
            {
                case DatabaseType.VortexAPI:
                case DatabaseType.VortexHistorianAPI:
                    return !string.IsNullOrEmpty(ConnectionSettings.EncryptedToken);

                case DatabaseType.PostgreSQL:
                case DatabaseType.MySQL:
                case DatabaseType.Oracle:
                case DatabaseType.SqlServer:
                    // Para relacionais, verificar campos básicos ou connection string
                    if (!string.IsNullOrEmpty(ConnectionSettings.ConnectionString))
                        return true;

                    return !string.IsNullOrEmpty(ConnectionSettings.Host) &&
                           ConnectionSettings.Port > 0 &&
                           !string.IsNullOrEmpty(ConnectionSettings.DatabaseName) &&
                           !string.IsNullOrEmpty(ConnectionSettings.Username);

                default:
                    return false;
            }
        }
    }
}
