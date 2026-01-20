using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using VortexExcelAddIn.Services;
using VortexExcelAddIn.Views;

namespace VortexExcelAddIn
{
    public partial class ThisAddIn
    {
        private Microsoft.Office.Tools.CustomTaskPane _taskPane;
        private VortexTaskPane _taskPaneControl;
        private Domain.Interfaces.IAutoRefreshService _autoRefreshService;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                // Criar controle WPF
                _taskPaneControl = new VortexTaskPane();

                // Criar UserControl para hospedar o controle WPF via ElementHost
                var userControl = new System.Windows.Forms.UserControl();
                var elementHost = new System.Windows.Forms.Integration.ElementHost
                {
                    Dock = System.Windows.Forms.DockStyle.Fill,
                    Child = _taskPaneControl
                };
                userControl.Controls.Add(elementHost);

                // Criar TaskPane
                _taskPane = this.CustomTaskPanes.Add(userControl, "Vortex Data Plugin");
                _taskPane.Width = 450;
                _taskPane.Visible = false;

                // Inicializar serviço de auto-refresh
                InitializeAutoRefreshService();

                // Subscrever evento de abertura de workbook
                this.Application.WorkbookOpen += Application_WorkbookOpen;

                try
                {
                    LoggingService.Info("Vortex Excel Add-in iniciado com sucesso");
                }
                catch
                {
                    // Ignorar erros de logging
                }
            }
            catch (Exception ex)
            {
                var errorMsg = $"Erro ao iniciar Vortex Add-in:\n\n{ex.Message}\n\nStack Trace:\n{ex.StackTrace}";

                try
                {
                    LoggingService.Fatal("Erro fatal ao iniciar add-in", ex);
                }
                catch
                {
                    // Ignorar erros de logging
                }

                MessageBox.Show(errorMsg, "Erro Fatal", MessageBoxButtons.OK, MessageBoxIcon.Error);

                // Re-lançar a exceção para que o Excel saiba que falhou
                throw;
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                // Cleanup auto-refresh
                _autoRefreshService?.Dispose();

                LoggingService.Info("Vortex Excel Add-in encerrando");
                LoggingService.Flush();
                LoggingService.Shutdown();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao encerrar add-in: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Mostra o TaskPane
        /// </summary>
        public void ShowTaskPane()
        {
            if (_taskPane != null)
            {
                _taskPane.Visible = true;
                LoggingService.Debug("TaskPane exibido");
            }
        }

        /// <summary>
        /// Esconde o TaskPane
        /// </summary>
        public void HideTaskPane()
        {
            if (_taskPane != null)
            {
                _taskPane.Visible = false;
                LoggingService.Debug("TaskPane ocultado");
            }
        }

        /// <summary>
        /// Alterna visibilidade do TaskPane
        /// </summary>
        public void ToggleTaskPane()
        {
            if (_taskPane != null)
            {
                _taskPane.Visible = !_taskPane.Visible;
                LoggingService.Debug($"TaskPane alternado: {_taskPane.Visible}");
            }
        }

        /// <summary>
        /// Inicializa o serviço de auto-refresh com dependências.
        /// </summary>
        private void InitializeAutoRefreshService()
        {
            try
            {
                // Obter Dispatcher do controle WPF
                var dispatcher = _taskPaneControl.Dispatcher;

                // Obter ConfigViewModel do MainViewModel
                var mainViewModel = _taskPaneControl.DataContext as ViewModels.MainViewModel;
                var configViewModel = mainViewModel?.ConfigViewModel;

                if (configViewModel == null)
                {
                    LoggingService.Error("Não foi possível inicializar auto-refresh: ConfigViewModel não encontrado");
                    return;
                }

                // Criar serviço de timer
                var timerService = new Services.SystemTimerService();

                // Criar serviço de auto-refresh
                _autoRefreshService = new Services.AutoRefreshService(
                    timerService,
                    configViewModel,
                    dispatcher);

                // Subscrever eventos para atualizações do ribbon
                _autoRefreshService.RefreshStarted += OnAutoRefreshStateChanged;
                _autoRefreshService.RefreshCompleted += OnAutoRefreshStateChanged;
                _autoRefreshService.RefreshFailed += OnAutoRefreshStateChanged;

                // Inicializar ViewModel no task pane
                var autoRefreshViewModel = new ViewModels.AutoRefreshViewModel(
                    _autoRefreshService,
                    mainViewModel.QueryViewModel);

                // Armazenar no MainViewModel para acesso
                mainViewModel.AutoRefreshViewModel = autoRefreshViewModel;

                // Subscrever ao evento de exportação de dados para habilitar botão Refresh
                mainViewModel.DataExportedToExcel += OnDataExportedToExcel;

                LoggingService.Info("Serviço de auto-refresh inicializado");
            }
            catch (Exception ex)
            {
                LoggingService.Error("Falha ao inicializar serviço de auto-refresh", ex);
            }
        }

        /// <summary>
        /// Handler para evento de abertura de workbook.
        /// </summary>
        private void Application_WorkbookOpen(Excel.Workbook wb)
        {
            try
            {
                LoggingService.Info($"Workbook aberto: {wb.Name}");

                // REMOVIDO: Não ativar automaticamente ao abrir workbook
                // O usuário deve clicar explicitamente em "Iniciar Refresh" para ativar
                // _autoRefreshService?.LoadAndActivate();

                // Atualizar estado do botão ribbon
                UpdateRibbonAutoRefreshButton();
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao processar abertura de workbook", ex);
            }
        }

        /// <summary>
        /// Exibe o diálogo de configuração de auto-refresh.
        /// Chamado pelo clique no botão do Ribbon.
        /// </summary>
        public void ShowAutoRefreshDialog()
        {
            try
            {
                var mainViewModel = _taskPaneControl.DataContext as ViewModels.MainViewModel;
                var autoRefreshViewModel = mainViewModel?.AutoRefreshViewModel;

                if (autoRefreshViewModel == null)
                {
                    MessageBox.Show("Auto-refresh não está disponível", "Erro",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var dialog = new Views.AutoRefreshDialog
                {
                    DataContext = autoRefreshViewModel
                };

                dialog.ShowDialog();

                // Atualizar ribbon após fechamento do diálogo
                UpdateRibbonAutoRefreshButton();
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao exibir diálogo de auto-refresh", ex);
                MessageBox.Show($"Erro: {ex.Message}", "Erro",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Exibe o diálogo de configuração de agregação temporal.
        /// Chamado pelo clique no botão "Tempo" do Ribbon.
        /// </summary>
        public void ShowTempoDialog()
        {
            try
            {
                var mainViewModel = _taskPaneControl.DataContext as ViewModels.MainViewModel;
                var configViewModel = mainViewModel?.ConfigViewModel;
                var queryViewModel = mainViewModel?.QueryViewModel;

                if (configViewModel == null || queryViewModel == null)
                {
                    MessageBox.Show("ViewModels não estão disponíveis", "Erro",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoggingService.Warn("ShowTempoDialog: ViewModels não disponíveis");
                    return;
                }

                // Criar ViewModel do diálogo
                var tempoViewModel = new ViewModels.TempoViewModel(
                    configViewModel,
                    queryViewModel);

                // Criar e exibir diálogo
                var dialog = new Views.TempoDialog
                {
                    DataContext = tempoViewModel
                };

                dialog.ShowDialog();

                LoggingService.Info("TempoDialog fechado");
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao exibir diálogo Tempo", ex);
                MessageBox.Show($"Erro ao abrir diálogo de agregação: {ex.Message}", "Erro",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Para o auto-refresh.
        /// Chamado pelo menu dropdown do botão do Ribbon.
        /// </summary>
        public void StopAutoRefresh()
        {
            try
            {
                _autoRefreshService?.Stop();
                UpdateRibbonAutoRefreshButton();
                LoggingService.Info("Auto-refresh parado via Ribbon menu");
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao parar auto-refresh", ex);
                MessageBox.Show($"Erro ao parar auto-refresh: {ex.Message}", "Erro",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Executa um refresh manual imediatamente.
        /// Chamado pelo menu dropdown do botão do Ribbon.
        /// </summary>
        public async void RefreshNow()
        {
            try
            {
                if (_autoRefreshService == null)
                {
                    MessageBox.Show("Serviço de refresh não está disponível", "Erro",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                LoggingService.Info("Refresh manual iniciado via Ribbon menu");
                await _autoRefreshService.RefreshNowAsync();
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao executar refresh manual", ex);
                MessageBox.Show($"Erro ao executar refresh: {ex.Message}", "Erro",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Handler para evento de exportação de dados.
        /// Habilita o botão Refresh quando dados são exportados para Excel.
        /// </summary>
        private void OnDataExportedToExcel(object sender, EventArgs e)
        {
            try
            {
                Globals.Ribbons.Ribbon1.EnableRefreshButton(true);
                LoggingService.Debug("Botão Refresh habilitado após exportação de dados");
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao habilitar botão Refresh", ex);
            }
        }

        /// <summary>
        /// Atualiza o botão do ribbon baseado no estado de auto-refresh.
        /// </summary>
        private void UpdateRibbonAutoRefreshButton()
        {
            try
            {
                var isActive = _autoRefreshService?.IsActive ?? false;
                Globals.Ribbons.Ribbon1.UpdateAutoRefreshButton(isActive);
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro ao atualizar botão do ribbon", ex);
            }
        }

        /// <summary>
        /// Handler para mudanças de estado do auto-refresh.
        /// </summary>
        private void OnAutoRefreshStateChanged(object sender, EventArgs e)
        {
            UpdateRibbonAutoRefreshButton();
        }

        #region Código gerado por VSTO

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
