using System.Collections.Generic;
using System.Xml.Serialization;

namespace VortexExcelAddIn.Models
{
    /// <summary>
    /// Configurações de conexão com campos dinâmicos para suportar múltiplos bancos de dados.
    /// </summary>
    public class DatabaseConnectionSettings
    {
        // Campos comuns a todos os bancos relacionais
        /// <summary>
        /// Endereço do servidor (hostname ou IP).
        /// </summary>
        public string Host { get; set; }

        /// <summary>
        /// Porta do servidor.
        /// </summary>
        public int Port { get; set; }

        /// <summary>
        /// Nome de usuário para autenticação.
        /// </summary>
        public string Username { get; set; }

        /// <summary>
        /// Senha criptografada com DPAPI.
        /// </summary>
        [XmlElement("EncryptedPassword")]
        public string EncryptedPassword { get; set; }

        // Campos específicos do InfluxDB
        /// <summary>
        /// URL completa do servidor InfluxDB (ex: http://localhost:8086).
        /// </summary>
        public string Url { get; set; }

        /// <summary>
        /// Token de autenticação do InfluxDB (criptografado com DPAPI).
        /// </summary>
        [XmlElement("EncryptedToken")]
        public string EncryptedToken { get; set; }

        /// <summary>
        /// Organização do InfluxDB.
        /// </summary>
        public string Org { get; set; }

        /// <summary>
        /// Bucket do InfluxDB (equivalente a database).
        /// </summary>
        public string Bucket { get; set; }

        // Campos para bancos relacionais
        /// <summary>
        /// Nome do banco de dados.
        /// </summary>
        public string DatabaseName { get; set; }

        /// <summary>
        /// Indica se deve usar SSL/TLS para conexão.
        /// </summary>
        public bool UseSsl { get; set; }

        /// <summary>
        /// String de conexão customizada (opcional, sobrescreve outros campos).
        /// </summary>
        public string ConnectionString { get; set; }

        /// <summary>
        /// Campos customizados adicionais (não serializados em XML).
        /// </summary>
        [XmlIgnore]
        public Dictionary<string, string> CustomFields { get; set; }

        public DatabaseConnectionSettings()
        {
            CustomFields = new Dictionary<string, string>();
            UseSsl = false;
            Port = 0;
        }
    }
}
