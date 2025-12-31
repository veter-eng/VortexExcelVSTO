using System;

namespace VortexExcelAddIn.Domain.Interfaces
{
    /// <summary>
    /// Abstração para funcionalidade de timer, permitindo testabilidade.
    /// Segue SOLID: ISP (Interface Segregation Principle) - interface mínima e focada.
    /// </summary>
    public interface ITimerService : IDisposable
    {
        /// <summary>
        /// Evento disparado quando o intervalo do timer decorre.
        /// </summary>
        event EventHandler Elapsed;

        /// <summary>
        /// Obtém ou define o intervalo em milissegundos.
        /// </summary>
        double Interval { get; set; }

        /// <summary>
        /// Obtém se o timer está habilitado.
        /// </summary>
        bool Enabled { get; }

        /// <summary>
        /// Inicia o timer.
        /// </summary>
        void Start();

        /// <summary>
        /// Para o timer.
        /// </summary>
        void Stop();
    }
}
