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

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                // Mostrar mensagem de início (DEBUG)
                MessageBox.Show("Vortex Add-in: Iniciando...", "Debug", MessageBoxButtons.OK, MessageBoxIcon.Information);

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

                try
                {
                    LoggingService.Info("Vortex Excel Add-in iniciado com sucesso");
                }
                catch
                {
                    // Ignorar erros de logging
                }

                // Mostrar mensagem de sucesso (DEBUG)
                MessageBox.Show("Vortex Add-in: Carregado com sucesso!\n\nProcure pela aba 'Suplementos' no Ribbon.", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
