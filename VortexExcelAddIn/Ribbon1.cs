using System;
using Microsoft.Office.Tools.Ribbon;

namespace VortexExcelAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            btnToggleTaskPane.Image = Properties.Resources.VortexIcon;
            btnAutoRefresh.Image = Properties.Resources.RefreshIcon;
            btnAutoRefresh.Enabled = false; // Desabilitado até dados serem exportados
            menuRefreshNow.Image = Properties.Resources.RefreshIcon;
            menuRefreshNow.Enabled = false; // Inicialmente desabilitado
            menuStopAutoRefresh.Image = Properties.Resources.StopIcon;
            menuStopAutoRefresh.Enabled = false; // Inicialmente desabilitado

            // Carregar ícone do botão Tempo
            try
            {
                btnTempo.Image = Properties.Resources.HourglassIcon;
            }
            catch
            {
                // Se o ícone não existir, deixa sem ícone (só texto)
            }
        }

        private void btnToggleTaskPane_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ToggleTaskPane();
        }

        private void btnAutoRefresh_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ShowAutoRefreshDialog();
        }

        private void menuStopAutoRefresh_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.StopAutoRefresh();
        }

        private void menuRefreshNow_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.RefreshNow();
        }

        private void btnTempo_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ShowTempoDialog();
        }

        /// <summary>
        /// Habilita ou desabilita o botão de Refresh.
        /// Chamado quando dados são exportados para Excel.
        /// </summary>
        public void EnableRefreshButton(bool enabled)
        {
            btnAutoRefresh.Enabled = enabled;
        }

        /// <summary>
        /// Atualiza o botão de auto-refresh baseado no estado ativo.
        /// Chamado por ThisAddIn quando o estado muda.
        /// </summary>
        public void UpdateAutoRefreshButton(bool isActive)
        {
            if (isActive)
            {
                btnAutoRefresh.Label = "Refresh (Ativo)";
                btnAutoRefresh.Image = Properties.Resources.RefreshActiveIcon;
                menuRefreshNow.Enabled = true;
                menuStopAutoRefresh.Enabled = true;
            }
            else
            {
                btnAutoRefresh.Label = "Refresh";
                btnAutoRefresh.Image = Properties.Resources.RefreshIcon;
                menuRefreshNow.Enabled = false;
                menuStopAutoRefresh.Enabled = false;
            }
        }
    }
}
