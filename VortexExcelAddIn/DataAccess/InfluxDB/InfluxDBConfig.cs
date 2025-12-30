namespace VortexExcelAddIn.DataAccess.InfluxDB
{
    /// <summary>
    /// Configuração de conexão com o InfluxDB.
    /// Movido de Models/ para DataAccess/InfluxDB/ como parte da refatoração SOLID.
    /// </summary>
    public class InfluxDBConfig
    {
        public string Url { get; set; }
        public string Token { get; set; }
        public string Org { get; set; }
        public string Bucket { get; set; }

        public InfluxDBConfig()
        {
            Url = string.Empty;
            Token = string.Empty;
            Org = string.Empty;
            Bucket = string.Empty;
        }

        public InfluxDBConfig(string url, string token, string org, string bucket)
        {
            Url = url;
            Token = token;
            Org = org;
            Bucket = bucket;
        }
    }
}
