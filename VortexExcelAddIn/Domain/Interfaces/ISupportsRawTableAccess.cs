using System.Collections.Generic;
using System.Threading.Tasks;
using VortexExcelAddIn.Models;

namespace VortexExcelAddIn.Domain.Interfaces
{
    /// <summary>
    /// Interface segregada para bancos relacionais com acesso direto a tabelas.
    /// Implementa ISP (Interface Segregation Principle) - apenas bancos relacionais implementam esta interface.
    /// Permite descoberta de schemas, tabelas e estrutura de dados.
    /// </summary>
    public interface ISupportsRawTableAccess
    {
        /// <summary>
        /// Lista todos os schemas disponíveis no banco de dados.
        /// </summary>
        /// <returns>Lista de nomes de schemas</returns>
        Task<List<string>> GetAvailableSchemasAsync();

        /// <summary>
        /// Lista todas as tabelas disponíveis em um schema específico.
        /// </summary>
        /// <param name="schemaName">Nome do schema</param>
        /// <returns>Lista de nomes de tabelas</returns>
        Task<List<string>> GetTablesInSchemaAsync(string schemaName);

        /// <summary>
        /// Obtém a estrutura (schema) de uma tabela específica.
        /// </summary>
        /// <param name="tableName">Nome da tabela</param>
        /// <param name="schemaName">Nome do schema (opcional, usa schema padrão se não especificado)</param>
        /// <returns>Objeto TableSchema com informações da estrutura da tabela</returns>
        Task<TableSchema> GetTableSchemaAsync(string tableName, string schemaName = null);
    }
}
