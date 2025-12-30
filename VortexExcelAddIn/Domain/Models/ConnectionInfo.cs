namespace VortexExcelAddIn.Domain.Models
{
    /// <summary>
    /// Informações sobre uma conexão ativa com banco de dados.
    /// </summary>
    public class ConnectionInfo
    {
        /// <summary>
        /// Tipo de banco de dados conectado.
        /// </summary>
        public DatabaseType DatabaseType { get; set; }

        /// <summary>
        /// Endereço do servidor (host:port ou URL).
        /// </summary>
        public string Host { get; set; }

        /// <summary>
        /// Nome do banco de dados ou bucket (InfluxDB).
        /// </summary>
        public string DatabaseName { get; set; }

        /// <summary>
        /// Nome de usuário conectado (se aplicável).
        /// </summary>
        public string Username { get; set; }

        /// <summary>
        /// Indica se a conexão está usando SSL/TLS.
        /// </summary>
        public bool IsSecure { get; set; }

        /// <summary>
        /// Versão do servidor (se disponível).
        /// </summary>
        public string ServerVersion { get; set; }

        /// <summary>
        /// Retorna uma representação em string das informações da conexão.
        /// </summary>
        /// <returns>String formatada com as informações da conexão</returns>
        public override string ToString()
        {
            return $"{DatabaseType.GetDisplayName()} - {Host}/{DatabaseName}";
        }
    }
}
