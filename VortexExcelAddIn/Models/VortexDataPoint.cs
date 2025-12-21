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
