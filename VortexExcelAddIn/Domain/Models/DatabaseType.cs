namespace VortexExcelAddIn.Domain.Models
{
    /// <summary>
    /// Tipos de bancos de dados suportados pelo sistema.
    /// </summary>
    public enum DatabaseType
    {
        /// <summary>
        /// PostgreSQL - Banco de dados relacional open source
        /// </summary>
        PostgreSQL,

        /// <summary>
        /// MySQL - Banco de dados relacional open source
        /// </summary>
        MySQL,

        /// <summary>
        /// Oracle Database - Banco de dados relacional empresarial
        /// </summary>
        Oracle,

        /// <summary>
        /// Microsoft SQL Server - Banco de dados relacional empresarial
        /// </summary>
        SqlServer,

        /// <summary>
        /// Vortex API - Acesso via API do Vortex IO (independente de banco)
        /// </summary>
        VortexAPI,

        /// <summary>
        /// Vortex Historian API - Acesso via API aos dados raw do Historian (dados_rabbitmq)
        /// </summary>
        VortexHistorianAPI
    }

    /// <summary>
    /// Métodos de extensão para DatabaseType.
    /// Fornece funcionalidades auxiliares sem violar o OCP (Open/Closed Principle).
    /// </summary>
    public static class DatabaseTypeExtensions
    {
        /// <summary>
        /// Retorna o nome de exibição formatado do tipo de banco de dados.
        /// </summary>
        /// <param name="type">Tipo de banco de dados</param>
        /// <returns>Nome formatado para exibição na UI</returns>
        public static string GetDisplayName(this DatabaseType type)
        {
            return type switch
            {
                DatabaseType.PostgreSQL => "PostgreSQL",
                DatabaseType.MySQL => "MySQL",
                DatabaseType.Oracle => "Oracle Database",
                DatabaseType.SqlServer => "SQL Server",
                DatabaseType.VortexAPI => "Servidor VortexIO",
                DatabaseType.VortexHistorianAPI => "Servidor Vortex Historian (API)",
                _ => type.ToString()
            };
        }

        /// <summary>
        /// Verifica se o tipo de banco é relacional (SQL).
        /// </summary>
        /// <param name="type">Tipo de banco de dados</param>
        /// <returns>True se for banco relacional, False caso contrário</returns>
        public static bool IsRelational(this DatabaseType type)
        {
            return type != DatabaseType.VortexAPI && type != DatabaseType.VortexHistorianAPI;
        }

        /// <summary>
        /// Verifica se o tipo usa API (ao invés de conexão direta).
        /// </summary>
        /// <param name="type">Tipo de banco de dados</param>
        /// <returns>True se for API, False caso contrário</returns>
        public static bool IsApi(this DatabaseType type)
        {
            return type == DatabaseType.VortexAPI || type == DatabaseType.VortexHistorianAPI;
        }

        /// <summary>
        /// Retorna a porta padrão para o tipo de banco de dados.
        /// </summary>
        /// <param name="type">Tipo de banco de dados</param>
        /// <returns>Número da porta padrão</returns>
        public static int GetDefaultPort(this DatabaseType type)
        {
            return type switch
            {
                DatabaseType.PostgreSQL => 5432,
                DatabaseType.MySQL => 3306,
                DatabaseType.Oracle => 1521,
                DatabaseType.SqlServer => 1433,
                DatabaseType.VortexAPI => 8000,
                DatabaseType.VortexHistorianAPI => 8000,
                _ => 0
            };
        }
    }
}
