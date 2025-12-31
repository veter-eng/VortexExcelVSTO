using System;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Threading;
using VortexExcelAddIn.Domain.Interfaces;
using VortexExcelAddIn.Models;
using VortexExcelAddIn.ViewModels;
using Excel = Microsoft.Office.Interop.Excel;

namespace VortexExcelAddIn.Services
{
    /// <summary>
    /// Serviço core para gerenciamento de funcionalidade de auto-refresh.
    ///
    /// DECISÕES ARQUITETURAIS:
    /// 1. Usa abstração ITimerService para testabilidade
    /// 2. Marshaling de callbacks do timer para main thread via Dispatcher
    /// 3. Gerenciamento de estado thread-safe
    /// 4. Dispara eventos para atualizações de UI (Observer pattern)
    /// 5. Integra com QueryViewModel/ConfigService patterns existentes
    ///
    /// CONSIDERAÇÕES DE SEGURANÇA:
    /// 1. Sincronização via lock para estado do timer
    /// 2. Acesso seguro a COM objects do Excel via Dispatcher
    /// 3. Isolamento de exceções - erros não quebram o timer
    /// 4. Validação de settings antes da execução
    /// </summary>
    public class AutoRefreshService : IAutoRefreshService
    {
        private readonly ITimerService _timer;
        private readonly ConfigViewModel _configViewModel;
        private readonly Dispatcher _dispatcher;
        private readonly object _lock = new object();

        private AutoRefreshSettings _settings;
        private DateTime? _nextRefreshTime;

        // Eventos
        public event EventHandler RefreshStarted;
        public event EventHandler<RefreshCompletedEventArgs> RefreshCompleted;
        public event EventHandler<RefreshErrorEventArgs> RefreshFailed;
        public event EventHandler NextRefreshTimeChanged;

        // Propriedades
        public AutoRefreshSettings Settings
        {
            get
            {
                lock (_lock)
                {
                    return _settings;
                }
            }
        }

        public bool IsActive
        {
            get
            {
                lock (_lock)
                {
                    return _timer?.Enabled ?? false;
                }
            }
        }

        public DateTime? NextRefreshTime
        {
            get
            {
                lock (_lock)
                {
                    return _nextRefreshTime;
                }
            }
            private set
            {
                lock (_lock)
                {
                    _nextRefreshTime = value;
                }
                NextRefreshTimeChanged?.Invoke(this, EventArgs.Empty);
            }
        }

        /// <summary>
        /// Construtor com injeção de dependências.
        /// </summary>
        /// <param name="timerService">Abstração de timer (injetado para testabilidade).</param>
        /// <param name="configViewModel">ConfigViewModel para obter conexão de banco.</param>
        /// <param name="dispatcher">Dispatcher da UI para marshaling de threads.</param>
        public AutoRefreshService(
            ITimerService timerService,
            ConfigViewModel configViewModel,
            Dispatcher dispatcher)
        {
            _timer = timerService ?? throw new ArgumentNullException(nameof(timerService));
            _configViewModel = configViewModel ?? throw new ArgumentNullException(nameof(configViewModel));
            _dispatcher = dispatcher ?? throw new ArgumentNullException(nameof(dispatcher));

            _timer.Elapsed += OnTimerElapsed;
        }

        /// <summary>
        /// Inicia auto-refresh com as configurações fornecidas.
        /// </summary>
        public void Start(AutoRefreshSettings settings)
        {
            if (settings == null || !settings.IsValid())
            {
                throw new ArgumentException("Configurações de auto-refresh inválidas", nameof(settings));
            }

            lock (_lock)
            {
                // Para timer existente se estiver rodando
                if (_timer.Enabled)
                {
                    _timer.Stop();
                }

                _settings = settings;
                _timer.Interval = settings.IntervalMinutes * 60 * 1000; // Converter para milissegundos
                _timer.Start();

                // Calcular próximo horário de atualização
                NextRefreshTime = DateTime.Now.AddMinutes(settings.IntervalMinutes);

                LoggingService.Info($"Auto-refresh iniciado: intervalo={settings.IntervalMinutes}min, target={settings.TargetSheetName}");
            }
        }

        /// <summary>
        /// Para a atualização automática.
        /// </summary>
        public void Stop()
        {
            lock (_lock)
            {
                _timer.Stop();
                _settings = null;
                NextRefreshTime = null;

                LoggingService.Info("Auto-refresh parado");
            }
        }

        /// <summary>
        /// Executa uma atualização manual imediatamente.
        /// </summary>
        public async Task RefreshNowAsync()
        {
            await ExecuteRefreshAsync();
        }

        /// <summary>
        /// Carrega configurações e ativa se habilitado.
        /// </summary>
        public void LoadAndActivate()
        {
            try
            {
                var settings = AutoRefreshConfigService.LoadSettings();

                if (settings != null && settings.IsEnabled)
                {
                    Start(settings);
                    LoggingService.Info("Auto-refresh ativado a partir de configurações salvas");
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error("Falha ao carregar e ativar auto-refresh", ex);
            }
        }

        /// <summary>
        /// Callback do timer - marshaling para main thread para acesso seguro ao Excel COM.
        /// </summary>
        private void OnTimerElapsed(object sender, EventArgs e)
        {
            // Marshal para UI thread para segurança do Excel COM
            _dispatcher.BeginInvoke(new Action(async () =>
            {
                await ExecuteRefreshAsync();
            }));
        }

        /// <summary>
        /// Lógica core de execução de refresh.
        /// Thread-safe, com isolamento de erros, event-driven.
        /// </summary>
        private async Task ExecuteRefreshAsync()
        {
            var startTime = DateTime.Now;
            AutoRefreshSettings currentSettings;

            lock (_lock)
            {
                currentSettings = _settings;
            }

            if (currentSettings == null)
                return;

            try
            {
                // Disparar evento de início
                RefreshStarted?.Invoke(this, EventArgs.Empty);

                LoggingService.Info("Execução de auto-refresh iniciada");

                // Obter conexão de banco do ConfigViewModel
                var connection = _configViewModel.GetConnection();
                if (connection == null)
                {
                    throw new InvalidOperationException("Nenhuma conexão de banco de dados configurada");
                }

                // Executar query usando parâmetros salvos
                var data = await connection.QueryDataAsync(currentSettings.QueryParameters);

                // Atualizar planilha Excel
                if (currentSettings.TargetSheetName == null)
                {
                    // Usar planilha ativa (Planilha Atual)
                    var activeSheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
                    if (activeSheet != null)
                    {
                        ExcelService.ExportToSheet(data, activeSheet.Name);
                    }
                    else
                    {
                        // Fallback: criar nova planilha se não houver planilha ativa
                        var sheetName = $"AutoRefresh_{DateTime.Now:yyyyMMdd_HHmmss}";
                        ExcelService.ExportToSheet(data, sheetName);
                    }
                }
                else if (!string.IsNullOrEmpty(currentSettings.TargetSheetName))
                {
                    // Atualizar planilha específica
                    ExcelService.ExportToSheet(data, currentSettings.TargetSheetName);
                }
                else
                {
                    // Criar nova planilha com timestamp (string vazia)
                    var sheetName = $"AutoRefresh_{DateTime.Now:yyyyMMdd_HHmmss}";
                    ExcelService.ExportToSheet(data, sheetName);
                }

                // Atualizar settings com horário da última atualização
                lock (_lock)
                {
                    if (_settings != null)
                    {
                        _settings.LastRefreshTime = DateTime.Now;
                        AutoRefreshConfigService.SaveSettings(_settings);
                    }

                    // Calcular próximo horário de atualização
                    if (_timer.Enabled)
                    {
                        NextRefreshTime = DateTime.Now.AddMinutes(currentSettings.IntervalMinutes);
                    }
                }

                // Disparar evento de sucesso
                var duration = DateTime.Now - startTime;
                RefreshCompleted?.Invoke(this, new RefreshCompletedEventArgs
                {
                    RecordsUpdated = data.Count,
                    RefreshTime = DateTime.Now,
                    Duration = duration
                });

                LoggingService.Info($"Auto-refresh concluído: {data.Count} registros em {duration.TotalSeconds:F1}s");
            }
            catch (Exception ex)
            {
                LoggingService.Error("Auto-refresh falhou", ex);

                // Disparar evento de erro
                RefreshFailed?.Invoke(this, new RefreshErrorEventArgs
                {
                    Error = ex,
                    ErrorTime = DateTime.Now
                });

                // Não para o timer em caso de erro - apenas loga e continua
            }
        }

        /// <summary>
        /// Libera recursos utilizados pelo serviço.
        /// </summary>
        public void Dispose()
        {
            Stop();
            _timer?.Dispose();
        }
    }
}
