using System;
using System.Timers;
using VortexExcelAddIn.Domain.Interfaces;

namespace VortexExcelAddIn.Services
{
    /// <summary>
    /// Implementação de ITimerService usando System.Timers.Timer.
    /// Thread-safe e apropriado para marshaling COM do Excel.
    /// </summary>
    public class SystemTimerService : ITimerService
    {
        private readonly Timer _timer;
        private readonly object _lock = new object();

        /// <summary>
        /// Evento disparado quando o intervalo do timer decorre.
        /// </summary>
        public event EventHandler Elapsed;

        /// <summary>
        /// Obtém ou define o intervalo em milissegundos.
        /// </summary>
        public double Interval
        {
            get
            {
                lock (_lock)
                {
                    return _timer.Interval;
                }
            }
            set
            {
                lock (_lock)
                {
                    _timer.Interval = value;
                }
            }
        }

        /// <summary>
        /// Obtém se o timer está habilitado.
        /// </summary>
        public bool Enabled
        {
            get
            {
                lock (_lock)
                {
                    return _timer.Enabled;
                }
            }
        }

        /// <summary>
        /// Inicializa uma nova instância de SystemTimerService.
        /// </summary>
        public SystemTimerService()
        {
            _timer = new Timer();
            _timer.AutoReset = true; // Repetir indefinidamente
            _timer.Elapsed += OnTimerElapsed;
        }

        /// <summary>
        /// Handler interno do evento Elapsed do timer.
        /// </summary>
        private void OnTimerElapsed(object sender, ElapsedEventArgs e)
        {
            // Propagar evento para subscribers
            Elapsed?.Invoke(this, EventArgs.Empty);
        }

        /// <summary>
        /// Inicia o timer.
        /// </summary>
        public void Start()
        {
            lock (_lock)
            {
                _timer.Start();
                LoggingService.Debug($"Timer iniciado com intervalo: {_timer.Interval}ms");
            }
        }

        /// <summary>
        /// Para o timer.
        /// </summary>
        public void Stop()
        {
            lock (_lock)
            {
                _timer.Stop();
                LoggingService.Debug("Timer parado");
            }
        }

        /// <summary>
        /// Libera recursos utilizados pelo timer.
        /// </summary>
        public void Dispose()
        {
            lock (_lock)
            {
                _timer?.Dispose();
            }
        }
    }
}
