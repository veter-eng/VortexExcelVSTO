using System;
using System.IO;
using System.Xml.Serialization;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using VortexExcelAddIn.Models;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace VortexExcelAddIn.Services
{
    /// <summary>
    /// Serviço para gerenciar configurações do add-in usando Custom XML Parts
    /// Port do ConfigService.ts
    /// </summary>
    public static class ConfigService
    {
        private const string ConfigNamespace = "http://vortex.com/influxdb-config";
        private const string ConfigRootElement = "InfluxDBConfig";

        /// <summary>
        /// Salva a configuração do InfluxDB no workbook atual
        /// </summary>
        public static void SaveConfig(InfluxDBConfig config)
        {
            if (config == null)
                throw new ArgumentNullException(nameof(config));

            try
            {
                var workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                if (workbook == null)
                {
                    LoggingService.Warn("Nenhum workbook ativo para salvar configuração");
                    return;
                }

                // Serializar config para XML
                var xmlContent = SerializeToXml(config);

                // Remover Custom XML Part existente
                var existing = GetCustomXmlPart(workbook);
                if (existing != null)
                {
                    existing.Delete();
                }

                // Adicionar novo Custom XML Part
                workbook.CustomXMLParts.Add(xmlContent);

                LoggingService.Info("Configuração salva no workbook");
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao salvar configuração", ex);
                throw;
            }
        }

        /// <summary>
        /// Carrega a configuração do InfluxDB do workbook atual
        /// </summary>
        public static InfluxDBConfig LoadConfig()
        {
            try
            {
                var workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                if (workbook == null)
                {
                    LoggingService.Warn("Nenhum workbook ativo para carregar configuração");
                    return GetDefaultConfig();
                }

                var customXmlPart = GetCustomXmlPart(workbook);
                if (customXmlPart == null)
                {
                    LoggingService.Debug("Nenhuma configuração encontrada, retornando padrão");
                    return GetDefaultConfig();
                }

                var xmlContent = customXmlPart.XML;
                var config = DeserializeFromXml(xmlContent);

                LoggingService.Info("Configuração carregada do workbook");
                return config;
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao carregar configuração", ex);
                return GetDefaultConfig();
            }
        }

        /// <summary>
        /// Limpa a configuração do workbook atual
        /// </summary>
        public static void ClearConfig()
        {
            try
            {
                var workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                if (workbook == null)
                {
                    LoggingService.Warn("Nenhum workbook ativo para limpar configuração");
                    return;
                }

                var customXmlPart = GetCustomXmlPart(workbook);
                if (customXmlPart != null)
                {
                    customXmlPart.Delete();
                    LoggingService.Info("Configuração limpa do workbook");
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao limpar configuração", ex);
                throw;
            }
        }

        /// <summary>
        /// Retorna a configuração padrão
        /// </summary>
        public static InfluxDBConfig GetDefaultConfig()
        {
            return new InfluxDBConfig
            {
                Url = "http://localhost:8086",
                Token = "",
                Org = "vortex",
                Bucket = "vortex_bucket"
            };
        }

        /// <summary>
        /// Verifica se existe configuração salva no workbook
        /// </summary>
        public static bool HasConfig()
        {
            try
            {
                var workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                if (workbook == null)
                    return false;

                return GetCustomXmlPart(workbook) != null;
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao verificar existência de configuração", ex);
                return false;
            }
        }

        #region Private Methods

        /// <summary>
        /// Obtém o Custom XML Part da configuração
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
                LoggingService.Error("Erro ao buscar Custom XML Part", ex);
                return null;
            }
        }

        /// <summary>
        /// Serializa a configuração para XML
        /// </summary>
        private static string SerializeToXml(InfluxDBConfig config)
        {
            try
            {
                var serializer = new XmlSerializer(typeof(InfluxDBConfig));
                using (var stringWriter = new StringWriter())
                {
                    var xmlNamespaces = new XmlSerializerNamespaces();
                    xmlNamespaces.Add("", ConfigNamespace);

                    serializer.Serialize(stringWriter, config, xmlNamespaces);
                    var xml = stringWriter.ToString();

                    // Adicionar namespace manualmente se necessário
                    if (!xml.Contains($"xmlns=\"{ConfigNamespace}\""))
                    {
                        xml = xml.Replace("<InfluxDBConfig>",
                            $"<InfluxDBConfig xmlns=\"{ConfigNamespace}\">");
                    }

                    return xml;
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao serializar configuração", ex);
                throw;
            }
        }

        /// <summary>
        /// Desserializa a configuração do XML
        /// </summary>
        private static InfluxDBConfig DeserializeFromXml(string xml)
        {
            try
            {
                var serializer = new XmlSerializer(typeof(InfluxDBConfig));
                using (var stringReader = new StringReader(xml))
                {
                    var config = (InfluxDBConfig)serializer.Deserialize(stringReader);
                    return config ?? GetDefaultConfig();
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao desserializar configuração", ex);
                return GetDefaultConfig();
            }
        }

        #endregion
    }
}
