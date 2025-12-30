using VortexExcelAddIn.Models;

namespace VortexExcelAddIn.DataAccess.PostgreSQL
{
    /// <summary>
    /// Configuração de conexão com o PostgreSQL.
    /// </summary>
    public class PostgreSqlConfig
    {
        public string Host { get; set; }
        public int Port { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
        public string DatabaseName { get; set; }
        public bool UseSsl { get; set; }
        public TableSchema TableSchema { get; set; }

        public PostgreSqlConfig()
        {
            Host = "localhost";
            Port = 5432;
            Username = string.Empty;
            Password = string.Empty;
            DatabaseName = string.Empty;
            UseSsl = false;
            TableSchema = new TableSchema();
        }

        /// <summary>
        /// Constrói a connection string para o Npgsql.
        /// </summary>
        public string BuildConnectionString()
        {
            var sslMode = UseSsl ? "Require" : "Prefer";
            return $"Host={Host};Port={Port};Database={DatabaseName};Username={Username};Password={Password};SSL Mode={sslMode};Timeout=30;Command Timeout=60";
        }
    }
}
