using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using VortexExcelAddIn.Domain.Interfaces;
using VortexExcelAddIn.Domain.Models;
using VortexExcelAddIn.Models;
using VortexExcelAddIn.Services;

namespace VortexExcelAddIn.Application.Strategies
{
    /// <summary>
    /// Strategy for applying REAL aggregation to Historian raw data.
    /// Uses ISupportsAggregation interface to aggregate on-the-fly using Flux queries.
    ///
    /// SOLID Principles Applied:
    /// - SRP: Single responsibility - handles only Historian aggregation logic
    /// - OCP: Can be extended without modifying existing code
    /// - LSP: Fully substitutable for IAggregationStrategy
    /// - DIP: Depends on IDataSourceConnection abstraction
    ///
    /// Behavior:
    /// For each selected aggregation type × each selected time window:
    ///   1. Execute QueryAggregatedDataAsync on ISupportsAggregation
    ///   2. Tag returned data points with metadata (AggregationType, TimeWindow)
    ///   3. Combine all results
    ///
    /// Special Cases:
    /// - MinMax requires TWO queries (min + max)
    /// - FirstLast requires TWO queries (first + last)
    /// - Delta requires TWO queries (first + last) then calculates difference
    /// </summary>
    public class HistorianAggregationStrategy : IAggregationStrategy
    {
        private readonly IDataSourceConnection _connection;

        public HistorianAggregationStrategy(IDataSourceConnection connection)
        {
            _connection = connection ?? throw new ArgumentNullException(nameof(connection));

            // Validate that connection supports aggregation
            if (!(_connection is ISupportsAggregation))
            {
                throw new ArgumentException(
                    "Connection must implement ISupportsAggregation to use HistorianAggregationStrategy",
                    nameof(connection));
            }
        }

        public async Task<List<VortexDataPoint>> ApplyAggregationAsync(
            QueryParams baseParams,
            AggregationConfiguration config)
        {
            var aggregator = _connection as ISupportsAggregation;
            var results = new List<VortexDataPoint>();

            LoggingService.Info($"[HistorianAggregationStrategy] Starting aggregation with {config.AggregationTypes.Count} types and {config.TimeWindows.Count} windows");

            // Nested loops: for each aggregation type, for each time window
            foreach (var aggType in config.AggregationTypes)
            {
                foreach (var window in config.TimeWindows)
                {
                    var windowPeriod = window.ToVortexIOFormat(); // "5m", "60m", etc.
                    var typeStr = aggType.ToVortexIOFormat();

                    LoggingService.Info($"[HistorianAggregationStrategy] Processing {typeStr} with window {windowPeriod}");

                    try
                    {
                        // Handle special cases that require multiple queries
                        if (aggType == VortexAggregationType.MinMax)
                        {
                            var minMaxData = await ProcessMinMaxAsync(aggregator, baseParams, windowPeriod, typeStr);
                            results.AddRange(minMaxData);
                        }
                        else if (aggType == VortexAggregationType.FirstLast)
                        {
                            var firstLastData = await ProcessFirstLastAsync(aggregator, baseParams, windowPeriod, typeStr);
                            results.AddRange(firstLastData);
                        }
                        else if (aggType == VortexAggregationType.Delta)
                        {
                            var deltaData = await ProcessDeltaAsync(aggregator, baseParams, windowPeriod, typeStr);
                            results.AddRange(deltaData);
                        }
                        else
                        {
                            // Simple aggregation types (Average, Total)
                            var mappedType = MapToAggregationType(aggType);
                            var data = await aggregator.QueryAggregatedDataAsync(baseParams, mappedType, windowPeriod);

                            // Tag with metadata
                            foreach (var point in data)
                            {
                                point.AggregationType = typeStr;
                                point.TimeWindow = windowPeriod;
                            }

                            results.AddRange(data);
                            LoggingService.Info($"[HistorianAggregationStrategy] {typeStr} with {windowPeriod}: {data.Count} points");
                        }
                    }
                    catch (Exception ex)
                    {
                        LoggingService.Error($"[HistorianAggregationStrategy] Error processing {typeStr} with {windowPeriod}", ex);
                        // Continue with other aggregations even if one fails
                    }
                }
            }

            LoggingService.Info($"[HistorianAggregationStrategy] Completed aggregation. Total points: {results.Count}");
            return results;
        }

        public string GetDescription()
        {
            return "Aplicar agregação aos dados brutos";
        }

        /// <summary>
        /// Processes MinMax aggregation by executing both Min and Max queries.
        /// </summary>
        private async Task<List<VortexDataPoint>> ProcessMinMaxAsync(
            ISupportsAggregation aggregator,
            QueryParams baseParams,
            string windowPeriod,
            string typeStr)
        {
            var results = new List<VortexDataPoint>();

            // Query Min
            var minData = await aggregator.QueryAggregatedDataAsync(baseParams, AggregationType.Min, windowPeriod);
            foreach (var point in minData)
            {
                point.AggregationType = $"{typeStr}_min";
                point.TimeWindow = windowPeriod;
            }
            results.AddRange(minData);

            // Query Max
            var maxData = await aggregator.QueryAggregatedDataAsync(baseParams, AggregationType.Max, windowPeriod);
            foreach (var point in maxData)
            {
                point.AggregationType = $"{typeStr}_max";
                point.TimeWindow = windowPeriod;
            }
            results.AddRange(maxData);

            LoggingService.Info($"[HistorianAggregationStrategy] MinMax with {windowPeriod}: {results.Count} points (min + max)");
            return results;
        }

        /// <summary>
        /// Processes FirstLast aggregation by executing both First and Last queries.
        /// </summary>
        private async Task<List<VortexDataPoint>> ProcessFirstLastAsync(
            ISupportsAggregation aggregator,
            QueryParams baseParams,
            string windowPeriod,
            string typeStr)
        {
            var results = new List<VortexDataPoint>();

            // Query First
            var firstData = await aggregator.QueryAggregatedDataAsync(baseParams, AggregationType.First, windowPeriod);
            foreach (var point in firstData)
            {
                point.AggregationType = $"{typeStr}_first";
                point.TimeWindow = windowPeriod;
            }
            results.AddRange(firstData);

            // Query Last
            var lastData = await aggregator.QueryAggregatedDataAsync(baseParams, AggregationType.Last, windowPeriod);
            foreach (var point in lastData)
            {
                point.AggregationType = $"{typeStr}_last";
                point.TimeWindow = windowPeriod;
            }
            results.AddRange(lastData);

            LoggingService.Info($"[HistorianAggregationStrategy] FirstLast with {windowPeriod}: {results.Count} points (first + last)");
            return results;
        }

        /// <summary>
        /// Processes Delta aggregation by querying First and Last, then calculating difference.
        /// Delta = Last - First
        /// </summary>
        private async Task<List<VortexDataPoint>> ProcessDeltaAsync(
            ISupportsAggregation aggregator,
            QueryParams baseParams,
            string windowPeriod,
            string typeStr)
        {
            // Query First and Last
            var firstData = await aggregator.QueryAggregatedDataAsync(baseParams, AggregationType.First, windowPeriod);
            var lastData = await aggregator.QueryAggregatedDataAsync(baseParams, AggregationType.Last, windowPeriod);

            // Group by Tag ID to match first with last
            var firstByTag = firstData.ToDictionary(p => p.TagId, p => p);
            var lastByTag = lastData.ToDictionary(p => p.TagId, p => p);

            var results = new List<VortexDataPoint>();

            foreach (var tagId in firstByTag.Keys)
            {
                if (lastByTag.ContainsKey(tagId))
                {
                    var first = firstByTag[tagId];
                    var last = lastByTag[tagId];

                    // Calculate delta
                    if (double.TryParse(first.Valor, out double firstVal) &&
                        double.TryParse(last.Valor, out double lastVal))
                    {
                        var delta = lastVal - firstVal;

                        var deltaPoint = new VortexDataPoint
                        {
                            Time = last.Time, // Use last timestamp
                            ColetorId = last.ColetorId,
                            GatewayId = last.GatewayId,
                            EquipmentId = last.EquipmentId,
                            TagId = last.TagId,
                            Valor = delta.ToString("F2"),
                            AggregationType = typeStr,
                            TimeWindow = windowPeriod
                        };

                        results.Add(deltaPoint);
                    }
                }
            }

            LoggingService.Info($"[HistorianAggregationStrategy] Delta with {windowPeriod}: {results.Count} points");
            return results;
        }

        /// <summary>
        /// Maps VortexAggregationType to existing AggregationType enum.
        /// </summary>
        private AggregationType MapToAggregationType(VortexAggregationType type)
        {
            return type switch
            {
                VortexAggregationType.Average => AggregationType.Mean,
                VortexAggregationType.Total => AggregationType.Sum,
                VortexAggregationType.MinMax => AggregationType.Max, // Handled separately
                VortexAggregationType.FirstLast => AggregationType.First, // Handled separately
                VortexAggregationType.Delta => AggregationType.Last, // Handled separately
                _ => AggregationType.Mean
            };
        }
    }
}
