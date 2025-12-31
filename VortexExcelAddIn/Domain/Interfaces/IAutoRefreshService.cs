using System;
using System.Threading.Tasks;
using VortexExcelAddIn.Models;

namespace VortexExcelAddIn.Domain.Interfaces
{
    /// <summary>
    /// Interface para serviço de atualização automática de dados.
    /// Segue SOLID: DIP (Dependency Inversion Principle) para testabilidade.
    /// </summary>
    public interface IAutoRefreshService : IDisposable
    {
        /// <summary>
        /// Evento disparado quando uma atualização é iniciada.
        /// </summary>
        event EventHandler RefreshStarted;

        /// <summary>
        /// Evento disparado quando uma atualização é concluída com sucesso.
        /// </summary>
        event EventHandler<RefreshCompletedEventArgs> RefreshCompleted;

        /// <summary>
        /// Evento disparado quando uma atualização falha.
        /// </summary>
        event EventHandler<RefreshErrorEventArgs> RefreshFailed;

        /// <summary>
        /// Evento disparado quando o horário da próxima atualização muda.
        /// </summary>
        event EventHandler NextRefreshTimeChanged;

        /// <summary>
        /// Obtém as configurações atuais de atualização automática.
        /// </summary>
        AutoRefreshSettings Settings { get; }

        /// <summary>
        /// Obtém se a atualização automática está atualmente ativa.
        /// </summary>
        bool IsActive { get; }

        /// <summary>
        /// Obtém o horário da próxima atualização agendada (null se não ativo).
        /// </summary>
        DateTime? NextRefreshTime { get; }

        /// <summary>
        /// Inicia a atualização automática com as configurações fornecidas.
        /// </summary>
        /// <param name="settings">Configurações de atualização automática.</param>
        void Start(AutoRefreshSettings settings);

        /// <summary>
        /// Para a atualização automática.
        /// </summary>
        void Stop();

        /// <summary>
        /// Executa uma atualização manual imediatamente (não afeta o timer).
        /// </summary>
        Task RefreshNowAsync();

        /// <summary>
        /// Carrega configurações do workbook e ativa se habilitado.
        /// Chamado ao abrir o workbook.
        /// </summary>
        void LoadAndActivate();
    }

    /// <summary>
    /// Argumentos de evento para atualização concluída com sucesso.
    /// </summary>
    public class RefreshCompletedEventArgs : EventArgs
    {
        /// <summary>
        /// Número de registros atualizados.
        /// </summary>
        public int RecordsUpdated { get; set; }

        /// <summary>
        /// Horário em que a atualização foi concluída.
        /// </summary>
        public DateTime RefreshTime { get; set; }

        /// <summary>
        /// Duração da operação de atualização.
        /// </summary>
        public TimeSpan Duration { get; set; }
    }

    /// <summary>
    /// Argumentos de evento para erros durante atualização.
    /// </summary>
    public class RefreshErrorEventArgs : EventArgs
    {
        /// <summary>
        /// Exceção que causou o erro.
        /// </summary>
        public Exception Error { get; set; }

        /// <summary>
        /// Horário em que o erro ocorreu.
        /// </summary>
        public DateTime ErrorTime { get; set; }
    }
}
