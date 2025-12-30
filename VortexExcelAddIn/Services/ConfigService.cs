using System;
using System.IO;
using System.Xml.Serialization;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using VortexExcelAddIn.Application.Security;
using VortexExcelAddIn.Domain.Models;
using VortexExcelAddIn.Models;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace VortexExcelAddIn.Services
{
    /// <summary>
    /// Serviço para gerenciar configurações do add-in usando Custom XML Parts.
    /// Refatorado para suportar múltiplos bancos de dados com backward compatibility.
    /// </summary>
    public static class ConfigService
    {
        // Configurações legadas (v1 - apenas InfluxDB)
        private const string LegacyConfigNamespace = "http://vortex.com/influxdb-config";
        private const string LegacyConfigRootElement = "InfluxDBConfig";

        // Configurações v2 (multi-banco de dados)
        private const string ConfigNamespaceV2 = "http://vortex.com/database-config-v2";
        private const string ConfigRootElementV2 = "UnifiedDatabaseConfig";

        // Mantém compatibilidade temporária
        private const string ConfigNamespace = LegacyConfigNamespace;
        private const string ConfigRootElement = LegacyConfigRootElement;

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
                Bucket = "vortex_data"
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

        #region V2 Methods - Multi-Database Support

        /// <summary>
        /// Salva a configuração unificada (v2) no workbook atual com criptografia.
        /// </summary>
        public static void SaveConfigV2(UnifiedDatabaseConfig config)
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

                // Criptografar credenciais antes de salvar
                var encryptor = new DPAPICredentialEncryptor();
                EncryptCredentials(config, encryptor);

                // Serializar config para XML
                var xmlContent = SerializeToXmlV2(config);

                // Remover Custom XML Part v2 existente
                var existingV2 = GetCustomXmlPartV2(workbook);
                if (existingV2 != null)
                {
                    existingV2.Delete();
                }

                // Remover configuração legada se existir (migração completa)
                var existingLegacy = GetCustomXmlPart(workbook);
                if (existingLegacy != null)
                {
                    existingLegacy.Delete();
                    LoggingService.Info("Configuração legada removida após migração");
                }

                // Adicionar novo Custom XML Part v2
                workbook.CustomXMLParts.Add(xmlContent);

                LoggingService.Info($"Configuração v2 salva no workbook (Tipo: {config.DatabaseType})");
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao salvar configuração v2", ex);
                throw;
            }
        }

        /// <summary>
        /// Carrega a configuração unificada (v2) do workbook atual com migração automática.
        /// </summary>
        public static UnifiedDatabaseConfig LoadConfigV2()
        {
            try
            {
                var workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                if (workbook == null)
                {
                    LoggingService.Warn("Nenhum workbook ativo para carregar configuração");
                    return GetDefaultConfigV2(DatabaseType.InfluxDB);
                }

                // 1. Tentar carregar configuração v2
                var customXmlPartV2 = GetCustomXmlPartV2(workbook);
                if (customXmlPartV2 != null)
                {
                    var xmlContent = customXmlPartV2.XML;
                    var config = DeserializeFromXmlV2(xmlContent);

                    // Descriptografar credenciais
                    var encryptor = new DPAPICredentialEncryptor();
                    DecryptCredentials(config, encryptor);

                    LoggingService.Info($"Configuração v2 carregada do workbook (Tipo: {config.DatabaseType})");
                    return config;
                }

                // 2. Se não encontrou v2, tentar migrar de v1 (configuração legada)
                LoggingService.Info("Configuração v2 não encontrada, tentando migrar configuração legada...");
                var migratedConfig = MigrateFromLegacyConfig();

                if (migratedConfig != null)
                {
                    LoggingService.Info("Configuração migrada com sucesso de v1 para v2");
                    return migratedConfig;
                }

                // 3. Retornar configuração padrão
                LoggingService.Debug("Nenhuma configuração encontrada, retornando padrão");
                return GetDefaultConfigV2(DatabaseType.InfluxDB);
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao carregar configuração v2", ex);
                return GetDefaultConfigV2(DatabaseType.InfluxDB);
            }
        }

        /// <summary>
        /// Migra configuração legada (v1 - InfluxDB) para nova configuração (v2 - UnifiedDatabaseConfig).
        /// </summary>
        public static UnifiedDatabaseConfig MigrateFromLegacyConfig()
        {
            try
            {
                var workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                if (workbook == null)
                {
                    LoggingService.Debug("Nenhum workbook ativo para migrar configuração");
                    return null;
                }

                // Buscar Custom XML Part legado
                var legacyPart = GetCustomXmlPart(workbook);
                if (legacyPart == null)
                {
                    LoggingService.Debug("Nenhuma configuração legada encontrada");
                    return null;
                }

                // Desserializar configuração legada
                var xmlContent = legacyPart.XML;
                var legacyConfig = DeserializeFromXml(xmlContent);

                if (legacyConfig == null)
                {
                    LoggingService.Warn("Falha ao desserializar configuração legada");
                    return null;
                }

                LoggingService.Info("Migrando configuração legada do InfluxDB para v2...");

                // Converter para novo formato
                var encryptor = new DPAPICredentialEncryptor();
                var unifiedConfig = new UnifiedDatabaseConfig
                {
                    DatabaseType = DatabaseType.InfluxDB,
                    ConfigVersion = 2,
                    ConnectionSettings = new DatabaseConnectionSettings
                    {
                        Url = legacyConfig.Url,
                        // Criptografar token durante migração
                        EncryptedToken = encryptor.Encrypt(legacyConfig.Token),
                        Org = legacyConfig.Org,
                        Bucket = legacyConfig.Bucket
                    }
                };

                // Salvar nova configuração (isso também remove a legada)
                SaveConfigV2(unifiedConfig);

                LoggingService.Info("Migração concluída com sucesso");
                return unifiedConfig;
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao migrar configuração legada", ex);
                return null;
            }
        }

        /// <summary>
        /// Retorna configuração padrão v2 para um tipo de banco de dados.
        /// </summary>
        public static UnifiedDatabaseConfig GetDefaultConfigV2(DatabaseType databaseType)
        {
            var factory = new Application.Factories.DatabaseConnectionFactory();
            return factory.CreateDefaultConfig(databaseType);
        }

        /// <summary>
        /// Verifica se existe configuração v2 salva no workbook.
        /// </summary>
        public static bool HasConfigV2()
        {
            try
            {
                var workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                if (workbook == null)
                    return false;

                return GetCustomXmlPartV2(workbook) != null;
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao verificar existência de configuração v2", ex);
                return false;
            }
        }

        /// <summary>
        /// Limpa a configuração v2 do workbook atual.
        /// </summary>
        public static void ClearConfigV2()
        {
            try
            {
                var workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                if (workbook == null)
                {
                    LoggingService.Warn("Nenhum workbook ativo para limpar configuração");
                    return;
                }

                var customXmlPartV2 = GetCustomXmlPartV2(workbook);
                if (customXmlPartV2 != null)
                {
                    customXmlPartV2.Delete();
                    LoggingService.Info("Configuração v2 limpa do workbook");
                }

                // Também limpar configuração legada se existir
                var customXmlPartLegacy = GetCustomXmlPart(workbook);
                if (customXmlPartLegacy != null)
                {
                    customXmlPartLegacy.Delete();
                    LoggingService.Info("Configuração legada limpa do workbook");
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao limpar configuração v2", ex);
                throw;
            }
        }

        #endregion

        #region V2 Private Methods

        /// <summary>
        /// Obtém o Custom XML Part da configuração v2.
        /// </summary>
        private static CustomXMLPart GetCustomXmlPartV2(Workbook workbook)
        {
            try
            {
                foreach (CustomXMLPart part in workbook.CustomXMLParts)
                {
                    if (part.NamespaceURI == ConfigNamespaceV2)
                    {
                        return part;
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao buscar Custom XML Part v2", ex);
                return null;
            }
        }

        /// <summary>
        /// Serializa a configuração v2 para XML.
        /// </summary>
        private static string SerializeToXmlV2(UnifiedDatabaseConfig config)
        {
            try
            {
                var serializer = new XmlSerializer(typeof(UnifiedDatabaseConfig));
                using (var stringWriter = new StringWriter())
                {
                    var xmlNamespaces = new XmlSerializerNamespaces();
                    xmlNamespaces.Add("", ConfigNamespaceV2);

                    serializer.Serialize(stringWriter, config, xmlNamespaces);
                    var xml = stringWriter.ToString();

                    // Adicionar namespace manualmente se necessário
                    if (!xml.Contains($"xmlns=\"{ConfigNamespaceV2}\""))
                    {
                        xml = xml.Replace("<UnifiedDatabaseConfig>",
                            $"<UnifiedDatabaseConfig xmlns=\"{ConfigNamespaceV2}\">");
                    }

                    return xml;
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao serializar configuração v2", ex);
                throw;
            }
        }

        /// <summary>
        /// Desserializa a configuração v2 do XML.
        /// </summary>
        private static UnifiedDatabaseConfig DeserializeFromXmlV2(string xml)
        {
            try
            {
                var serializer = new XmlSerializer(typeof(UnifiedDatabaseConfig));
                using (var stringReader = new StringReader(xml))
                {
                    var config = (UnifiedDatabaseConfig)serializer.Deserialize(stringReader);
                    return config ?? GetDefaultConfigV2(DatabaseType.InfluxDB);
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao desserializar configuração v2", ex);
                return GetDefaultConfigV2(DatabaseType.InfluxDB);
            }
        }

        /// <summary>
        /// Criptografa as credenciais na configuração.
        /// </summary>
        private static void EncryptCredentials(UnifiedDatabaseConfig config, DPAPICredentialEncryptor encryptor)
        {
            if (config?.ConnectionSettings == null)
                return;

            try
            {
                // Criptografar baseado no tipo de banco
                if (config.DatabaseType == DatabaseType.InfluxDB)
                {
                    // Criptografar token do InfluxDB
                    if (!string.IsNullOrEmpty(config.ConnectionSettings.EncryptedToken))
                    {
                        config.ConnectionSettings.EncryptedToken =
                            encryptor.Encrypt(config.ConnectionSettings.EncryptedToken);
                    }
                }
                else
                {
                    // Criptografar senha para bancos relacionais
                    if (!string.IsNullOrEmpty(config.ConnectionSettings.EncryptedPassword))
                    {
                        config.ConnectionSettings.EncryptedPassword =
                            encryptor.Encrypt(config.ConnectionSettings.EncryptedPassword);
                    }
                }

                LoggingService.Debug("Credenciais criptografadas com sucesso");
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao criptografar credenciais", ex);
                throw;
            }
        }

        /// <summary>
        /// Descriptografa as credenciais na configuração.
        /// </summary>
        private static void DecryptCredentials(UnifiedDatabaseConfig config, DPAPICredentialEncryptor encryptor)
        {
            if (config?.ConnectionSettings == null)
                return;

            try
            {
                // Descriptografar baseado no tipo de banco
                if (config.DatabaseType == DatabaseType.InfluxDB)
                {
                    // Descriptografar token do InfluxDB
                    if (!string.IsNullOrEmpty(config.ConnectionSettings.EncryptedToken))
                    {
                        config.ConnectionSettings.EncryptedToken =
                            encryptor.Decrypt(config.ConnectionSettings.EncryptedToken);
                    }
                }
                else
                {
                    // Descriptografar senha para bancos relacionais
                    if (!string.IsNullOrEmpty(config.ConnectionSettings.EncryptedPassword))
                    {
                        config.ConnectionSettings.EncryptedPassword =
                            encryptor.Decrypt(config.ConnectionSettings.EncryptedPassword);
                    }
                }

                LoggingService.Debug("Credenciais descriptografadas com sucesso");
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao descriptografar credenciais", ex);
                throw;
            }
        }

        #endregion
    }
}
