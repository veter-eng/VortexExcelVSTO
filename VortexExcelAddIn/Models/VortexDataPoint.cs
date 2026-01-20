using System;

namespace VortexExcelAddIn.Models
{
    /// <summary>
    /// Representa um ponto de dados do sistema Vortex importado do InfluxDB
    /// </summary>
    public class VortexDataPoint
    {
        public DateTime Time { get; set; }
        public string ColetorId { get; set; }
        public string GatewayId { get; set; }
        public string EquipmentId { get; set; }
        public string TagId { get; set; }
        public string Valor { get; set; }

        /// <summary>
        /// Tipo de agregação aplicada a este ponto de dados (opcional).
        /// Exemplos: "average", "total", "min_max", "first_last", "delta"
        /// Usado para identificar agregações em queries com múltiplos tipos.
        /// </summary>
        public string AggregationType { get; set; }

        /// <summary>
        /// Janela de tempo da agregação (opcional).
        /// Exemplos: "5m", "15m", "30m", "60m"
        /// Usado para identificar janelas de tempo em queries com múltiplas janelas.
        /// </summary>
        public string TimeWindow { get; set; }

        public VortexDataPoint()
        {
            Time = DateTime.UtcNow;
            ColetorId = string.Empty;
            GatewayId = string.Empty;
            EquipmentId = string.Empty;
            TagId = string.Empty;
            Valor = string.Empty;
        }

        public VortexDataPoint(DateTime time, string coletorId, string gatewayId,
            string equipmentId, string tagId, string valor)
        {
            Time = time;
            ColetorId = coletorId;
            GatewayId = gatewayId;
            EquipmentId = equipmentId;
            TagId = tagId;
            Valor = valor;
        }
    }
}
