using System.Collections.Generic;
using System.Threading.Tasks;
using VortexExcelAddIn.Domain.Models;
using VortexExcelAddIn.Models;

namespace VortexExcelAddIn.Domain.Interfaces
{
    /// <summary>
    /// Interface segregada para bancos de dados que suportam agregação de dados.
    /// Implementa ISP (Interface Segregation Principle) - clientes não são forçados a depender de métodos que não usam.
    /// Apenas bancos que suportam agregação (InfluxDB, etc.) implementam esta interface.
    /// </summary>
    public interface ISupportsAggregation
    {
        /// <summary>
        /// Executa consulta com agregação de dados.
        /// </summary>
        /// <param name="parameters">Parâmetros de consulta base</param>
        /// <param name="aggregation">Tipo de agregação (Mean, Min, Max, Count, Sum, etc.)</param>
        /// <param name="windowPeriod">Período da janela de agregação (ex: "1m", "1h", "1d")</param>
        /// <returns>Lista de pontos de dados agregados</returns>
        Task<List<VortexDataPoint>> QueryAggregatedDataAsync(
            QueryParams parameters,
            AggregationType aggregation,
            string windowPeriod = "1m");
    }
}
