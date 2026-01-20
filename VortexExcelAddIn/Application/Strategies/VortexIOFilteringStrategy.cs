using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using VortexExcelAddIn.Domain.Interfaces;
using VortexExcelAddIn.Models;
using VortexExcelAddIn.Services;

namespace VortexExcelAddIn.Application.Strategies
{
    /// <summary>
    /// Strategy for FILTERING pre-aggregated VortexIO data.
    /// Does NOT re-aggregate - VortexIO data is already aggregated by Airflow backend.
    ///
    /// SOLID Principles Applied:
    /// - SRP: Single responsibility - handles only VortexIO filtering logic
    /// - OCP: Can be extended without modifying existing code
    /// - LSP: Fully substitutable for IAggregationStrategy
    /// - DIP: Depends on IDataSourceConnection abstraction
    ///
    /// Behavior:
    /// 1. Query ALL pre-aggregated data from dados_airflow (single query)
    /// 2. Filter results locally by parsing equipment_id field
    /// 3. Keep only data matching selected aggregation types and time windows
    ///
    /// VortexIO Data Format:
    /// - gateway_id: Contains _field name (e.g., "avg_valor", "sum_valor", "min_valor")
    /// - equipment_id: Contains aggregation_type (e.g., "average_60m", "total_30m", "min_max_5m")
    ///
    /// Parsing Logic:
    /// Split equipment_id by '_' to extract:
    /// - Aggregation type (e.g., "average" from "average_60m")
    /// - Time window (e.g., "60m" from "average_60m")
    /// </summary>
    public class VortexIOFilteringStrategy : IAggregationStrategy
    {
        private readonly IDataSourceConnection _connection;

        public VortexIOFilteringStrategy(IDataSourceConnection connection)
        {
            _connection = connection ?? throw new ArgumentNullException(nameof(connection));
        }

        public async Task<List<VortexDataPoint>> ApplyAggregationAsync(
            QueryParams baseParams,
            AggregationConfiguration config)
        {
            LoggingService.Info($"[VortexIOFilteringStrategy] Starting filtering with {config.AggregationTypes.Count} types and {config.TimeWindows.Count} windows");

            try
            {
                // IMPORTANT: VortexIO uses equipment_id and gateway_id fields differently:
                // - equipment_id contains aggregation metadata (e.g., "average_60m")
                // - gateway_id contains field names (e.g., "avg_valor")
                // We must clear these filters to get all pre-aggregated data, then filter locally.
                var vortexIOParams = new QueryParams
                {
                    ColetorId = baseParams.ColetorId, // Keep coletor filter
                    TagId = baseParams.TagId,         // Keep tag filter
                    GatewayId = null,                 // Clear - contains field names in VortexIO
                    EquipmentId = null,               // Clear - contains aggregation metadata in VortexIO
                    StartTime = baseParams.StartTime,
                    EndTime = baseParams.EndTime,
                    Limit = baseParams.Limit
                };

                LoggingService.Info($"[VortexIOFilteringStrategy] VortexIO query params - ColetorId: {vortexIOParams.ColetorId}, TagId: {vortexIOParams.TagId}, Time: {vortexIOParams.StartTime:yyyy-MM-dd} to {vortexIOParams.EndTime:yyyy-MM-dd}");

                // Query ALL pre-aggregated data (single query, no aggregation params)
                var allData = await _connection.QueryDataAsync(vortexIOParams);
                LoggingService.Info($"[VortexIOFilteringStrategy] Retrieved {allData.Count} total pre-aggregated points");

                // DEBUG: Log sample data to understand structure
                if (allData.Count > 0)
                {
                    var samples = allData.Take(5).ToList();
                    foreach (var sample in samples)
                    {
                        LoggingService.Info($"[VortexIOFilteringStrategy] SAMPLE DATA - ColetorId: '{sample.ColetorId}', GatewayId: '{sample.GatewayId}', EquipmentId: '{sample.EquipmentId}', TagId: '{sample.TagId}', Valor: '{sample.Valor}'");
                    }

                    // Log unique equipment_id values to understand the format
                    var uniqueEquipmentIds = allData.Select(d => d.EquipmentId).Distinct().Take(20).ToList();
                    LoggingService.Info($"[VortexIOFilteringStrategy] UNIQUE EquipmentIds (first 20): {string.Join(", ", uniqueEquipmentIds)}");

                    var uniqueGatewayIds = allData.Select(d => d.GatewayId).Distinct().Take(20).ToList();
                    LoggingService.Info($"[VortexIOFilteringStrategy] UNIQUE GatewayIds (first 20): {string.Join(", ", uniqueGatewayIds)}");
                }

                // Build allowed filters
                // Use ToVortexIOFormats() to handle MinMax -> [min, max] and FirstLast -> [first, last]
                var allowedTypes = new HashSet<string>(
                    config.AggregationTypes.SelectMany(t => t.ToVortexIOFormats()));

                var allowedWindows = new HashSet<string>(
                    config.TimeWindows.Select(w => w.ToVortexIOFormat()));

                LoggingService.Info($"[VortexIOFilteringStrategy] Allowed types: {string.Join(", ", allowedTypes)}");
                LoggingService.Info($"[VortexIOFilteringStrategy] Allowed windows: {string.Join(", ", allowedWindows)}");

                // Filter by equipment_id field which contains aggregation metadata
                // Format is always "type_window" e.g., "average_60m", "min_60m", "max_60m"
                var filtered = allData.Where(point =>
                {
                    if (string.IsNullOrEmpty(point.EquipmentId))
                        return false;

                    // Parse "average_60m" format - find last underscore to split type and window
                    var lastUnderscoreIndex = point.EquipmentId.LastIndexOf('_');
                    if (lastUnderscoreIndex <= 0)
                    {
                        LoggingService.Debug($"[VortexIOFilteringStrategy] Skipping point with invalid format: {point.EquipmentId}");
                        return false;
                    }

                    var type = point.EquipmentId.Substring(0, lastUnderscoreIndex); // "average", "min", "max", etc.
                    var window = point.EquipmentId.Substring(lastUnderscoreIndex + 1); // "60m", "5m", etc.

                    bool matches = allowedTypes.Contains(type) && allowedWindows.Contains(window);

                    if (matches)
                    {
                        // Tag point with metadata for clarity
                        point.AggregationType = type;
                        point.TimeWindow = window;
                    }

                    return matches;
                }).ToList();

                LoggingService.Info($"[VortexIOFilteringStrategy] Filtering complete. {filtered.Count} points matched criteria (from {allData.Count} total)");
                return filtered;
            }
            catch (Exception ex)
            {
                LoggingService.Error("[VortexIOFilteringStrategy] Error during filtering", ex);
                throw; // Re-throw to let caller handle
            }
        }

        public string GetDescription()
        {
            return "Filtrar dados já agregados";
        }
    }
}
