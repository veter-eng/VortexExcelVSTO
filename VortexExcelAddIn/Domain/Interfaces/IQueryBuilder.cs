using VortexExcelAddIn.Models;

namespace VortexExcelAddIn.Domain.Interfaces
{
    /// <summary>
    /// Interface para construção de queries específicas de cada banco de dados.
    /// Implementa SRP (Single Responsibility Principle) - responsabilidade única de construir queries.
    /// Implementa OCP (Open/Closed Principle) - novas implementações podem ser adicionadas sem modificar código existente.
    /// </summary>
    public interface IQueryBuilder
    {
        /// <summary>
        /// Constrói query para buscar dados com filtros aplicados.
        /// </summary>
        /// <param name="parameters">Parâmetros de consulta (filtros, intervalo de tempo, limite)</param>
        /// <param name="schema">Schema da tabela (apenas para bancos relacionais, opcional para time-series)</param>
        /// <returns>String da query no formato específico do banco (SQL, Flux, etc.)</returns>
        string BuildDataQuery(QueryParams parameters, TableSchema schema = null);

        /// <summary>
        /// Constrói query simples para teste de conexão.
        /// </summary>
        /// <returns>Query de teste (ex: "SELECT 1" para SQL, query básica para InfluxDB)</returns>
        string BuildTestQuery();

        /// <summary>
        /// Valida se os parâmetros fornecidos são compatíveis com este query builder.
        /// </summary>
        /// <param name="parameters">Parâmetros a serem validados</param>
        /// <returns>True se válidos, False caso contrário</returns>
        bool ValidateParameters(QueryParams parameters);
    }
}
