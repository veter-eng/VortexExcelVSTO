namespace VortexExcelAddIn.DataAccess.VortexAPI
{
    /// <summary>
    /// Configuração para acesso à API Vortex IO com credenciais inline.
    /// A API é sempre http://localhost:8000 (valor fixo) e as credenciais do InfluxDB
    /// são enviadas diretamente na requisição ao invés de usar um ID de conexão gerenciado.
    /// </summary>
    public class VortexApiConfig
    {
        /// <summary>
        /// URL base da API VortexIO (fixo)
        /// </summary>
        public string ApiUrl => "http://localhost:8000";

        /// <summary>
        /// Host do InfluxDB (ex: localhost)
        /// </summary>
        public string InfluxHost { get; set; }

        /// <summary>
        /// Porta do InfluxDB (padrão: 8086)
        /// </summary>
        public int InfluxPort { get; set; }

        /// <summary>
        /// Organization do InfluxDB
        /// </summary>
        public string InfluxOrg { get; set; }

        /// <summary>
        /// Bucket do InfluxDB
        /// </summary>
        public string InfluxBucket { get; set; }

        /// <summary>
        /// Token de autenticação do InfluxDB
        /// </summary>
        public string InfluxToken { get; set; }

        /// <summary>
        /// Timeout para requisições HTTP em segundos (padrão: 30)
        /// </summary>
        public int Timeout { get; set; }

        public VortexApiConfig()
        {
            // Usar nome do container Docker ao invés de localhost
            // pois o backend está em Docker e precisa acessar o InfluxDB via rede Docker
            InfluxHost = "vortex_influxdb";
            InfluxPort = 8086;
            InfluxOrg = "vortex";
            InfluxBucket = "dados_airflow"; // Bucket com dados agregados do Airflow
            Timeout = 30;
        }

        /// <summary>
        /// Valida a configuração.
        /// </summary>
        /// <returns>True se a configuração é válida</returns>
        public bool IsValid()
        {
            return !string.IsNullOrWhiteSpace(InfluxHost) &&
                   InfluxPort > 0 &&
                   !string.IsNullOrWhiteSpace(InfluxOrg) &&
                   !string.IsNullOrWhiteSpace(InfluxBucket) &&
                   !string.IsNullOrWhiteSpace(InfluxToken) &&
                   Timeout > 0;
        }
    }
}
