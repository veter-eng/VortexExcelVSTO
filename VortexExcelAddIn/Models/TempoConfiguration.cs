using System.Collections.Generic;
using VortexExcelAddIn.Domain.Models;

namespace VortexExcelAddIn.Models
{
    /// <summary>
    /// Configuração persistente do diálogo Tempo.
    /// Mantém as últimas seleções do usuário entre aberturas do diálogo.
    /// </summary>
    public static class TempoConfiguration
    {
        /// <summary>
        /// Último servidor selecionado.
        /// </summary>
        public static DatabaseType? LastSelectedServer { get; set; }

        /// <summary>
        /// Últimos tipos de agregação selecionados.
        /// </summary>
        public static HashSet<VortexAggregationType> LastSelectedAggregationTypes { get; set; }
            = new HashSet<VortexAggregationType>();

        /// <summary>
        /// Últimas janelas de tempo selecionadas.
        /// </summary>
        public static HashSet<TimeWindow> LastSelectedTimeWindows { get; set; }
            = new HashSet<TimeWindow>();

        /// <summary>
        /// Limpa todas as seleções persistidas.
        /// </summary>
        public static void Clear()
        {
            LastSelectedServer = null;
            LastSelectedAggregationTypes.Clear();
            LastSelectedTimeWindows.Clear();
        }
    }
}
