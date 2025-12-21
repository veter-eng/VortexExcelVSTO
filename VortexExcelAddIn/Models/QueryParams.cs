using System;

namespace VortexExcelAddIn.Models
{
    /// <summary>
    /// Parâmetros para consulta de dados no InfluxDB
    /// </summary>
    public class QueryParams
    {
        public string ColetorId { get; set; }
        public string GatewayId { get; set; }
        public string EquipmentId { get; set; }
        public string TagId { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public int? Limit { get; set; }

        public QueryParams()
        {
            ColetorId = null;
            GatewayId = null;
            EquipmentId = null;
            TagId = null;
            StartTime = DateTime.UtcNow.AddHours(-24);
            EndTime = DateTime.UtcNow;
            Limit = 1000;
        }

        public QueryParams(DateTime startTime, DateTime endTime,
            string coletorId = null, string gatewayId = null,
            string equipmentId = null, string tagId = null, int? limit = 1000)
        {
            StartTime = startTime;
            EndTime = endTime;
            ColetorId = coletorId;
            GatewayId = gatewayId;
            EquipmentId = equipmentId;
            TagId = tagId;
            Limit = limit;
        }
    }
}
