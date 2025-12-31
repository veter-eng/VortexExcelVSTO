using System;
using System.IO;
using System.Xml.Serialization;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using VortexExcelAddIn.Models;

namespace VortexExcelAddIn.Services
{
    /// <summary>
    /// Serviço para persistência de configurações de auto-refresh usando Custom XML Parts.
    /// Segue o mesmo padrão do ConfigService existente.
    /// </summary>
    public static class AutoRefreshConfigService
    {
        private const string ConfigNamespace = "http://vortex.com/auto-refresh-v1";

        /// <summary>
        /// Salva as configurações de auto-refresh no workbook atual.
        /// </summary>
        /// <param name="settings">Configurações a serem salvas.</param>
        public static void SaveSettings(AutoRefreshSettings settings)
        {
            if (settings == null)
                throw new ArgumentNullException(nameof(settings));

            try
            {
                var workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                if (workbook == null)
                {
                    LoggingService.Warn("Nenhum workbook ativo para salvar configurações de auto-refresh");
                    return;
                }

                // Serializar configurações para XML
                var xmlContent = SerializeToXml(settings);

                // Remover Custom XML Part existente
                var existing = GetCustomXmlPart(workbook);
                if (existing != null)
                {
                    existing.Delete();
                }

                // Adicionar novo Custom XML Part
                workbook.CustomXMLParts.Add(xmlContent);

                LoggingService.Info("Configurações de auto-refresh salvas no workbook");
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao salvar configurações de auto-refresh", ex);
                throw;
            }
        }

        /// <summary>
        /// Carrega as configurações de auto-refresh do workbook atual.
        /// </summary>
        /// <returns>Configurações carregadas ou null se não encontradas.</returns>
        public static AutoRefreshSettings LoadSettings()
        {
            try
            {
                var workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                if (workbook == null)
                {
                    LoggingService.Debug("Nenhum workbook ativo para carregar configurações de auto-refresh");
                    return null;
                }

                var customXmlPart = GetCustomXmlPart(workbook);
                if (customXmlPart == null)
                {
                    LoggingService.Debug("Nenhuma configuração de auto-refresh encontrada");
                    return null;
                }

                var xmlContent = customXmlPart.XML;
                var settings = DeserializeFromXml(xmlContent);

                LoggingService.Info("Configurações de auto-refresh carregadas do workbook");
                return settings;
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao carregar configurações de auto-refresh", ex);
                return null;
            }
        }

        /// <summary>
        /// Limpa as configurações de auto-refresh do workbook atual.
        /// </summary>
        public static void ClearSettings()
        {
            try
            {
                var workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                if (workbook == null)
                {
                    LoggingService.Warn("Nenhum workbook ativo para limpar configurações de auto-refresh");
                    return;
                }

                var customXmlPart = GetCustomXmlPart(workbook);
                if (customXmlPart != null)
                {
                    customXmlPart.Delete();
                    LoggingService.Info("Configurações de auto-refresh limpas do workbook");
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao limpar configurações de auto-refresh", ex);
                throw;
            }
        }

        #region Private Methods

        /// <summary>
        /// Obtém o Custom XML Part das configurações de auto-refresh.
        /// </summary>
        private static CustomXMLPart GetCustomXmlPart(Workbook workbook)
        {
            try
            {
                foreach (CustomXMLPart part in workbook.CustomXMLParts)
                {
                    if (part.NamespaceURI == ConfigNamespace)
                    {
                        return part;
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao buscar Custom XML Part de auto-refresh", ex);
                return null;
            }
        }

        /// <summary>
        /// Serializa as configurações para XML.
        /// </summary>
        private static string SerializeToXml(AutoRefreshSettings settings)
        {
            try
            {
                var serializer = new XmlSerializer(typeof(AutoRefreshSettings));
                using (var stringWriter = new StringWriter())
                {
                    var xmlNamespaces = new XmlSerializerNamespaces();
                    xmlNamespaces.Add("", ConfigNamespace);

                    serializer.Serialize(stringWriter, settings, xmlNamespaces);
                    return stringWriter.ToString();
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao serializar configurações de auto-refresh", ex);
                throw;
            }
        }

        /// <summary>
        /// Desserializa as configurações do XML.
        /// </summary>
        private static AutoRefreshSettings DeserializeFromXml(string xml)
        {
            try
            {
                var serializer = new XmlSerializer(typeof(AutoRefreshSettings));
                using (var stringReader = new StringReader(xml))
                {
                    var settings = (AutoRefreshSettings)serializer.Deserialize(stringReader);
                    return settings;
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao desserializar configurações de auto-refresh", ex);
                return null;
            }
        }

        #endregion
    }
}
