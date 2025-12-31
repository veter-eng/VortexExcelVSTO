using System;
using System.Xml.Serialization;

namespace VortexExcelAddIn.Models
{
    /// <summary>
    /// Configurações para funcionalidade de atualização automática de dados.
    /// Serializável para persistência via CustomXMLParts.
    /// </summary>
    [XmlRoot("AutoRefreshSettings", Namespace = "http://vortex.com/auto-refresh-v1")]
    public class AutoRefreshSettings
    {
        /// <summary>
        /// Indica se a atualização automática está habilitada.
        /// </summary>
        public bool IsEnabled { get; set; }

        /// <summary>
        /// Intervalo de atualização em minutos (1-60).
        /// </summary>
        public int IntervalMinutes { get; set; }

        /// <summary>
        /// Nome da planilha alvo para atualização.
        /// Se null ou vazio, cria nova planilha a cada atualização.
        /// </summary>
        public string TargetSheetName { get; set; }

        /// <summary>
        /// Número máximo de resultados a buscar (fixo em 1000 conforme requisitos).
        /// </summary>
        public int ResultLimit { get; set; }

        /// <summary>
        /// Timestamp da última atualização bem-sucedida.
        /// Usado para exibição de status e diagnósticos.
        /// </summary>
        public DateTime? LastRefreshTime { get; set; }

        /// <summary>
        /// Parâmetros de consulta capturados do QueryViewModel.
        /// Usados para executar a atualização automática.
        /// </summary>
        public QueryParams QueryParameters { get; set; }

        /// <summary>
        /// Inicializa uma nova instância de AutoRefreshSettings com valores padrão.
        /// </summary>
        public AutoRefreshSettings()
        {
            IsEnabled = false;
            IntervalMinutes = 5; // Padrão: 5 minutos
            ResultLimit = 1000;
            TargetSheetName = string.Empty;
            QueryParameters = new QueryParams();
        }

        /// <summary>
        /// Valida as configurações quanto a consistência.
        /// </summary>
        /// <returns>True se as configurações são válidas, false caso contrário.</returns>
        public bool IsValid()
        {
            return IntervalMinutes >= 1 &&
                   IntervalMinutes <= 60 &&
                   ResultLimit > 0;
        }
    }
}
