using System.Collections.Generic;
using System.Threading.Tasks;
using VortexExcelAddIn.Models;

namespace VortexExcelAddIn.Domain.Interfaces
{
    /// <summary>
    /// Strategy interface for applying aggregation or filtering logic to data queries.
    /// Implements Strategy Pattern (GoF) to handle different server behaviors.
    ///
    /// SOLID Principles:
    /// - OCP (Open/Closed): New strategies can be added without modifying existing code
    /// - LSP (Liskov Substitution): All implementations are interchangeable
    /// - DIP (Dependency Inversion): High-level modules depend on this abstraction
    ///
    /// Implementations:
    /// - HistorianAggregationStrategy: Applies REAL aggregation to raw Historian data
    /// - VortexIOFilteringStrategy: FILTERS pre-aggregated VortexIO data
    /// </summary>
    public interface IAggregationStrategy
    {
        /// <summary>
        /// Applies aggregation or filtering based on strategy implementation.
        ///
        /// For Historian: Executes multiple aggregation queries using ISupportsAggregation.
        /// For VortexIO: Queries all data and filters locally by aggregation_type field.
        /// </summary>
        /// <param name="baseParams">Base query parameters (time range, filters, etc.)</param>
        /// <param name="config">Aggregation configuration with selected types and windows</param>
        /// <returns>List of data points matching aggregation configuration</returns>
        Task<List<VortexDataPoint>> ApplyAggregationAsync(
            QueryParams baseParams,
            AggregationConfiguration config);

        /// <summary>
        /// Gets user-friendly description of what this strategy does.
        /// Used for UI display to help users understand behavior.
        /// </summary>
        /// <returns>Description string in Portuguese</returns>
        string GetDescription();
    }
}
