using System;
using NLog;

namespace VortexExcelAddIn.Services
{
    /// <summary>
    /// Serviço centralizado de logging usando NLog
    /// </summary>
    public static class LoggingService
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        static LoggingService()
        {
            try
            {
                // Obter o diretório do assembly atual
                string assemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string assemblyDirectory = System.IO.Path.GetDirectoryName(assemblyPath);
                string configPath = System.IO.Path.Combine(assemblyDirectory, "NLog.config");

                // Carregar configuração do NLog.config se existir
                if (System.IO.File.Exists(configPath))
                {
                    LogManager.Setup().LoadConfigurationFromFile(configPath);
                }
            }
            catch
            {
                // Se falhar ao carregar a configuração, o NLog usará configuração padrão
                // Ignora o erro para não quebrar a inicialização do add-in
            }
        }

        /// <summary>
        /// Log de informação
        /// </summary>
        public static void Info(string message)
        {
            Logger.Info(message);
        }

        /// <summary>
        /// Log de informação com parâmetros
        /// </summary>
        public static void Info(string message, params object[] args)
        {
            Logger.Info(message, args);
        }

        /// <summary>
        /// Log de debug (apenas em modo debug)
        /// </summary>
        public static void Debug(string message)
        {
            Logger.Debug(message);
        }

        /// <summary>
        /// Log de debug com parâmetros
        /// </summary>
        public static void Debug(string message, params object[] args)
        {
            Logger.Debug(message, args);
        }

        /// <summary>
        /// Log de warning
        /// </summary>
        public static void Warn(string message)
        {
            Logger.Warn(message);
        }

        /// <summary>
        /// Log de warning com parâmetros
        /// </summary>
        public static void Warn(string message, params object[] args)
        {
            Logger.Warn(message, args);
        }

        /// <summary>
        /// Log de erro
        /// </summary>
        public static void Error(string message)
        {
            Logger.Error(message);
        }

        /// <summary>
        /// Log de erro com exceção
        /// </summary>
        public static void Error(string message, Exception ex)
        {
            Logger.Error(ex, message);
        }

        /// <summary>
        /// Log de erro com parâmetros
        /// </summary>
        public static void Error(Exception ex, string message, params object[] args)
        {
            Logger.Error(ex, message, args);
        }

        /// <summary>
        /// Log de erro fatal
        /// </summary>
        public static void Fatal(string message)
        {
            Logger.Fatal(message);
        }

        /// <summary>
        /// Log de erro fatal com exceção
        /// </summary>
        public static void Fatal(string message, Exception ex)
        {
            Logger.Fatal(ex, message);
        }

        /// <summary>
        /// Faz flush de todos os logs pendentes
        /// </summary>
        public static void Flush()
        {
            LogManager.Flush();
        }

        /// <summary>
        /// Libera recursos do NLog (chamar no shutdown)
        /// </summary>
        public static void Shutdown()
        {
            LogManager.Shutdown();
        }
    }
}
