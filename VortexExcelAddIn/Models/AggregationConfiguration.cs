using System.Collections.Generic;
using VortexExcelAddIn.Domain.Models;

namespace VortexExcelAddIn.Models
{
    /// <summary>
    /// Configuration for time-based aggregation.
    /// Immutable value object following DDD principles.
    /// Used to pass aggregation settings from UI to business logic.
    /// </summary>
    public class AggregationConfiguration
    {
        /// <summary>
        /// Selected aggregation types (can be multiple).
        /// Examples: Average, Total, MinMax
        /// </summary>
        public List<VortexAggregationType> AggregationTypes { get; set; }

        /// <summary>
        /// Selected time windows (can be multiple).
        /// Examples: 5min, 15min, 30min, 60min
        /// </summary>
        public List<TimeWindow> TimeWindows { get; set; }

        /// <summary>
        /// Server type context - determines aggregation vs filtering behavior.
        /// </summary>
        public DatabaseType ServerType { get; set; }

        /// <summary>
        /// Validates that configuration has at least one aggregation type and one time window.
        /// </summary>
        /// <returns>True if valid, false otherwise</returns>
        public bool IsValid() => AggregationTypes?.Count > 0 && TimeWindows?.Count > 0;
    }

    /// <summary>
    /// VortexIO aggregation types matching backend naming conventions.
    /// Maps to the aggregation types available in the VortexIO backend.
    /// </summary>
    public enum VortexAggregationType
    {
        /// <summary>
        /// Average aggregation - calculates mean value per tag.
        /// Maps to "average" in VortexIO backend.
        /// </summary>
        Average,

        /// <summary>
        /// Total (sum) aggregation - calculates sum of values per tag.
        /// Maps to "total" in VortexIO backend.
        /// </summary>
        Total,

        /// <summary>
        /// Min/Max aggregation - calculates both minimum and maximum values.
        /// Maps to "min_max" in VortexIO backend.
        /// </summary>
        MinMax,

        /// <summary>
        /// First/Last aggregation - captures first and last values in time window.
        /// Maps to "first_last" in VortexIO backend.
        /// </summary>
        FirstLast,

        /// <summary>
        /// Delta aggregation - calculates difference between last and first values.
        /// Maps to "delta" in VortexIO backend.
        /// </summary>
        Delta
    }

    /// <summary>
    /// Time window options for aggregation.
    /// Values represent minutes.
    /// </summary>
    public enum TimeWindow
    {
        /// <summary>
        /// 5-minute time window
        /// </summary>
        FiveMinutes = 5,

        /// <summary>
        /// 15-minute time window
        /// </summary>
        FifteenMinutes = 15,

        /// <summary>
        /// 30-minute time window
        /// </summary>
        ThirtyMinutes = 30,

        /// <summary>
        /// 60-minute (1 hour) time window
        /// </summary>
        SixtyMinutes = 60
    }

    /// <summary>
    /// Extension methods for aggregation configuration enums.
    /// Provides display names and format conversions.
    /// </summary>
    public static class AggregationConfigurationExtensions
    {
        /// <summary>
        /// Returns display name for aggregation type in Portuguese.
        /// </summary>
        public static string ToDisplayName(this VortexAggregationType type)
        {
            return type switch
            {
                VortexAggregationType.Average => "Média (Average)",
                VortexAggregationType.Total => "Total (Sum)",
                VortexAggregationType.MinMax => "Mínimo/Máximo (Min/Max)",
                VortexAggregationType.FirstLast => "Primeiro/Último (First/Last)",
                VortexAggregationType.Delta => "Delta",
                _ => type.ToString()
            };
        }

        /// <summary>
        /// Returns VortexIO backend format strings.
        /// Some types like MinMax and FirstLast map to multiple values (min, max) and (first, last).
        /// </summary>
        public static string ToVortexIOFormat(this VortexAggregationType type)
        {
            return type switch
            {
                VortexAggregationType.Average => "average",
                VortexAggregationType.Total => "total",
                VortexAggregationType.MinMax => "min_max", // Legacy - use ToVortexIOFormats() instead
                VortexAggregationType.FirstLast => "first_last", // Legacy - use ToVortexIOFormats() instead
                VortexAggregationType.Delta => "delta",
                _ => "average"
            };
        }

        /// <summary>
        /// Returns all VortexIO backend format strings for a type.
        /// MinMax returns both "min" and "max", FirstLast returns both "first" and "last".
        /// </summary>
        public static IEnumerable<string> ToVortexIOFormats(this VortexAggregationType type)
        {
            return type switch
            {
                VortexAggregationType.Average => new[] { "average" },
                VortexAggregationType.Total => new[] { "total" },
                VortexAggregationType.MinMax => new[] { "min", "max" },
                VortexAggregationType.FirstLast => new[] { "first", "last" },
                VortexAggregationType.Delta => new[] { "delta" },
                _ => new[] { "average" }
            };
        }

        /// <summary>
        /// Maps VortexAggregationType to existing AggregationType enum.
        /// Used for ISupportsAggregation interface compatibility.
        /// </summary>
        public static AggregationType ToAggregationType(this VortexAggregationType type)
        {
            return type switch
            {
                VortexAggregationType.Average => AggregationType.Mean,
                VortexAggregationType.Total => AggregationType.Sum,
                VortexAggregationType.MinMax => AggregationType.Max, // Will need separate Min query
                VortexAggregationType.FirstLast => AggregationType.First, // Will need separate Last query
                VortexAggregationType.Delta => AggregationType.Last, // Calculate delta from First and Last
                _ => AggregationType.Mean
            };
        }

        /// <summary>
        /// Returns display name for time window in Portuguese.
        /// </summary>
        public static string ToDisplayName(this TimeWindow window)
        {
            return window switch
            {
                TimeWindow.FiveMinutes => "5 minutos",
                TimeWindow.FifteenMinutes => "15 minutos",
                TimeWindow.ThirtyMinutes => "30 minutos",
                TimeWindow.SixtyMinutes => "60 minutos (1 hora)",
                _ => $"{(int)window} minutos"
            };
        }

        /// <summary>
        /// Returns VortexIO backend format for time window (e.g., "5m", "60m").
        /// </summary>
        public static string ToVortexIOFormat(this TimeWindow window)
        {
            return $"{(int)window}m";
        }
    }
}
