using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using VortexExcelAddIn.Domain.Models;
using VortexExcelAddIn.Models;

namespace VortexExcelAddIn.Domain.Interfaces
{
    /// <summary>
    /// Interface base para conexões com fontes de dados.
    /// Implementa os princípios SOLID: DIP (Dependency Inversion) e ISP (Interface Segregation).
    /// Todas as implementações de conexão de banco de dados devem implementar esta interface.
    /// </summary>
    public interface IDataSourceConnection : IDisposable
    {
        /// <summary>
        /// Testa a conectividade com a fonte de dados.
        /// </summary>
        /// <returns>Resultado do teste de conexão com informações de sucesso/falha</returns>
        Task<ConnectionResult> TestConnectionAsync();

        /// <summary>
        /// Executa consulta e retorna dados baseado nos parâmetros fornecidos.
        /// </summary>
        /// <param name="parameters">Parâmetros de consulta (filtros, intervalo de tempo, etc.)</param>
        /// <returns>Lista de pontos de dados</returns>
        Task<List<VortexDataPoint>> QueryDataAsync(QueryParams parameters);

        /// <summary>
        /// Retorna informações sobre a conexão atual.
        /// </summary>
        /// <returns>Informações da conexão (tipo, host, database, etc.)</returns>
        ConnectionInfo GetConnectionInfo();

        /// <summary>
        /// Tipo de banco de dados desta conexão.
        /// </summary>
        DatabaseType DatabaseType { get; }
    }
}
