using System;
using VortexExcelAddIn.Application.Strategies;
using VortexExcelAddIn.Domain.Interfaces;
using VortexExcelAddIn.Domain.Models;

namespace VortexExcelAddIn.Application.Factories
{
    /// <summary>
    /// Factory for creating aggregation strategies based on database type.
    /// Implements Factory Pattern and follows OCP (Open/Closed Principle).
    ///
    /// SOLID Principles Applied:
    /// - SRP: Single responsibility - creates strategies only
    /// - OCP: New database types can be added without modifying existing code
    /// - DIP: Returns IAggregationStrategy abstraction, not concrete types
    ///
    /// Strategy Selection Logic:
    /// - VortexHistorianAPI → HistorianAggregationStrategy (real-time aggregation)
    /// - VortexAPI → VortexIOFilteringStrategy (client-side filtering)
    /// - Other types → NotSupportedException
    ///
    /// Usage:
    /// var strategy = AggregationStrategyFactory.CreateStrategy(serverType, connection);
    /// var results = await strategy.ApplyAggregationAsync(params, config);
    /// </summary>
    public static class AggregationStrategyFactory
    {
        /// <summary>
        /// Creates appropriate aggregation strategy based on database type.
        /// </summary>
        /// <param name="serverType">Type of database server</param>
        /// <param name="connection">Data source connection implementing IDataSourceConnection</param>
        /// <returns>Strategy implementation matching server type</returns>
        /// <exception cref="ArgumentNullException">If connection is null</exception>
        /// <exception cref="NotSupportedException">If server type doesn't support aggregation</exception>
        public static IAggregationStrategy CreateStrategy(
            DatabaseType serverType,
            IDataSourceConnection connection)
        {
            if (connection == null)
            {
                throw new ArgumentNullException(nameof(connection),
                    "Connection cannot be null");
            }

            return serverType switch
            {
                DatabaseType.VortexHistorianAPI =>
                    new HistorianAggregationStrategy(connection),

                DatabaseType.VortexAPI =>
                    new VortexIOFilteringStrategy(connection),

                _ => throw new NotSupportedException(
                    $"Aggregation is not supported for database type: {serverType}. " +
                    $"Supported types: VortexHistorianAPI, VortexAPI")
            };
        }

        /// <summary>
        /// Checks if a database type supports aggregation.
        /// </summary>
        /// <param name="serverType">Type of database server to check</param>
        /// <returns>True if aggregation is supported, false otherwise</returns>
        public static bool IsAggregationSupported(DatabaseType serverType)
        {
            return serverType == DatabaseType.VortexHistorianAPI ||
                   serverType == DatabaseType.VortexAPI;
        }

        /// <summary>
        /// Gets description of aggregation behavior for a database type.
        /// </summary>
        /// <param name="serverType">Type of database server</param>
        /// <returns>Description string in Portuguese</returns>
        public static string GetAggregationDescription(DatabaseType serverType)
        {
            return serverType switch
            {
                DatabaseType.VortexHistorianAPI => "Aplicar agregação aos dados brutos",
                DatabaseType.VortexAPI => "Filtrar dados já agregados",
                _ => "Agregação não suportada"
            };
        }
    }
}
