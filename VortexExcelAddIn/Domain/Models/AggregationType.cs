namespace VortexExcelAddIn.Domain.Models
{
    /// <summary>
    /// Tipos de agregação suportados para consultas de dados.
    /// </summary>
    public enum AggregationType
    {
        /// <summary>
        /// Média dos valores.
        /// </summary>
        Mean,

        /// <summary>
        /// Valor mínimo.
        /// </summary>
        Min,

        /// <summary>
        /// Valor máximo.
        /// </summary>
        Max,

        /// <summary>
        /// Contagem de registros.
        /// </summary>
        Count,

        /// <summary>
        /// Soma dos valores.
        /// </summary>
        Sum,

        /// <summary>
        /// Desvio padrão.
        /// </summary>
        StdDev,

        /// <summary>
        /// Primeiro valor no intervalo.
        /// </summary>
        First,

        /// <summary>
        /// Último valor no intervalo.
        /// </summary>
        Last
    }

    /// <summary>
    /// Métodos de extensão para AggregationType.
    /// </summary>
    public static class AggregationTypeExtensions
    {
        /// <summary>
        /// Retorna o nome de exibição da agregação.
        /// </summary>
        /// <param name="type">Tipo de agregação</param>
        /// <returns>Nome formatado para exibição</returns>
        public static string GetDisplayName(this AggregationType type)
        {
            return type switch
            {
                AggregationType.Mean => "Média",
                AggregationType.Min => "Mínimo",
                AggregationType.Max => "Máximo",
                AggregationType.Count => "Contagem",
                AggregationType.Sum => "Soma",
                AggregationType.StdDev => "Desvio Padrão",
                AggregationType.First => "Primeiro",
                AggregationType.Last => "Último",
                _ => type.ToString()
            };
        }

        /// <summary>
        /// Retorna o nome da função em Flux (InfluxDB).
        /// </summary>
        /// <param name="type">Tipo de agregação</param>
        /// <returns>Nome da função Flux</returns>
        public static string ToFluxFunction(this AggregationType type)
        {
            return type switch
            {
                AggregationType.Mean => "mean",
                AggregationType.Min => "min",
                AggregationType.Max => "max",
                AggregationType.Count => "count",
                AggregationType.Sum => "sum",
                AggregationType.StdDev => "stddev",
                AggregationType.First => "first",
                AggregationType.Last => "last",
                _ => "mean"
            };
        }

        /// <summary>
        /// Retorna o nome da função SQL para bancos relacionais.
        /// </summary>
        /// <param name="type">Tipo de agregação</param>
        /// <returns>Nome da função SQL</returns>
        public static string ToSqlFunction(this AggregationType type)
        {
            return type switch
            {
                AggregationType.Mean => "AVG",
                AggregationType.Min => "MIN",
                AggregationType.Max => "MAX",
                AggregationType.Count => "COUNT",
                AggregationType.Sum => "SUM",
                AggregationType.StdDev => "STDDEV",
                AggregationType.First => "FIRST_VALUE",
                AggregationType.Last => "LAST_VALUE",
                _ => "AVG"
            };
        }
    }
}
