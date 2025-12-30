using System;
using System.Collections.Generic;

namespace VortexExcelAddIn.Domain.Models
{
    /// <summary>
    /// Resultado de uma tentativa de conexão com banco de dados.
    /// Fornece informações detalhadas sobre sucesso/falha da conexão.
    /// </summary>
    public class ConnectionResult
    {
        /// <summary>
        /// Indica se a conexão foi estabelecida com sucesso.
        /// </summary>
        public bool IsSuccessful { get; set; }

        /// <summary>
        /// Mensagem descritiva sobre o resultado da conexão.
        /// </summary>
        public string Message { get; set; }

        /// <summary>
        /// Exceção capturada em caso de falha (se houver).
        /// </summary>
        public Exception Exception { get; set; }

        /// <summary>
        /// Tempo de resposta da conexão (latência).
        /// </summary>
        public TimeSpan Latency { get; set; }

        /// <summary>
        /// Metadados adicionais sobre a conexão (versão do servidor, etc.).
        /// </summary>
        public Dictionary<string, object> Metadata { get; set; }

        public ConnectionResult()
        {
            Metadata = new Dictionary<string, object>();
        }

        /// <summary>
        /// Cria um resultado de conexão bem-sucedida.
        /// </summary>
        /// <param name="message">Mensagem de sucesso</param>
        /// <returns>ConnectionResult com IsSuccessful = true</returns>
        public static ConnectionResult Success(string message = "Conexão estabelecida com sucesso")
        {
            return new ConnectionResult
            {
                IsSuccessful = true,
                Message = message,
                Latency = TimeSpan.Zero
            };
        }

        /// <summary>
        /// Cria um resultado de conexão falha.
        /// </summary>
        /// <param name="message">Mensagem de erro</param>
        /// <param name="ex">Exceção capturada (opcional)</param>
        /// <returns>ConnectionResult com IsSuccessful = false</returns>
        public static ConnectionResult Failure(string message, Exception ex = null)
        {
            return new ConnectionResult
            {
                IsSuccessful = false,
                Message = message,
                Exception = ex
            };
        }
    }
}
