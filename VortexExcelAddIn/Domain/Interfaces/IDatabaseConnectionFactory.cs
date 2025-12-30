using VortexExcelAddIn.Domain.Models;
using VortexExcelAddIn.Models;

namespace VortexExcelAddIn.Domain.Interfaces
{
    /// <summary>
    /// Interface para factory de conexões de banco de dados.
    /// Implementa DIP (Dependency Inversion Principle) - permite injeção de dependência.
    /// </summary>
    public interface IDatabaseConnectionFactory
    {
        /// <summary>
        /// Cria uma conexão baseado na configuração fornecida.
        /// </summary>
        /// <param name="config">Configuração unificada do banco de dados</param>
        /// <returns>Conexão criada implementando IDataSourceConnection</returns>
        IDataSourceConnection CreateConnection(UnifiedDatabaseConfig config);

        /// <summary>
        /// Cria configuração padrão para um tipo de banco.
        /// </summary>
        /// <param name="type">Tipo de banco de dados</param>
        /// <returns>Configuração padrão para o tipo especificado</returns>
        UnifiedDatabaseConfig CreateDefaultConfig(DatabaseType type);

        /// <summary>
        /// Verifica se um tipo de banco é suportado.
        /// </summary>
        /// <param name="type">Tipo de banco de dados</param>
        /// <returns>True se suportado, False caso contrário</returns>
        bool IsSupported(DatabaseType type);
    }
}
