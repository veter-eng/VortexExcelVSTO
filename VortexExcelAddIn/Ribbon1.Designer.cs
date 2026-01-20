namespace VortexExcelAddIn
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnToggleTaskPane = this.Factory.CreateRibbonButton();
            this.btnAutoRefresh = this.Factory.CreateRibbonSplitButton();
            this.btnTempo = this.Factory.CreateRibbonButton();
            this.menuAutoRefresh = this.Factory.CreateRibbonMenu();
            this.menuRefreshNow = this.Factory.CreateRibbonButton();
            this.menuStopAutoRefresh = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            //
            // tab1
            //
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Custom;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "Vortex";
            this.tab1.Name = "tab1";
            //
            // group1
            //
            this.group1.Items.Add(this.btnToggleTaskPane);
            this.group1.Items.Add(this.btnAutoRefresh);
            this.group1.Items.Add(this.btnTempo);
            this.group1.Label = "Vortex Data";
            this.group1.Name = "group1";
            //
            // btnToggleTaskPane
            //
            this.btnToggleTaskPane.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnToggleTaskPane.Label = "Vortex Plugin";
            this.btnToggleTaskPane.Name = "btnToggleTaskPane";
            this.btnToggleTaskPane.ShowImage = true;
            this.btnToggleTaskPane.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnToggleTaskPane_Click);
            //
            // btnAutoRefresh
            //
            this.btnAutoRefresh.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAutoRefresh.Items.Add(this.menuRefreshNow);
            this.btnAutoRefresh.Items.Add(this.menuStopAutoRefresh);
            this.btnAutoRefresh.Label = "Refresh";
            this.btnAutoRefresh.Name = "btnAutoRefresh";
            this.btnAutoRefresh.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAutoRefresh_Click);
            //
            // btnTempo
            //
            this.btnTempo.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnTempo.Label = "Tempo";
            this.btnTempo.Name = "btnTempo";
            this.btnTempo.ShowImage = true;
            this.btnTempo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnTempo_Click);
            //
            // menuRefreshNow
            //
            this.menuRefreshNow.Label = "Atualizar Agora";
            this.menuRefreshNow.Name = "menuRefreshNow";
            this.menuRefreshNow.ShowImage = true;
            this.menuRefreshNow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.menuRefreshNow_Click);
            //
            // menuStopAutoRefresh
            //
            this.menuStopAutoRefresh.Label = "Parar Refresh";
            this.menuStopAutoRefresh.Name = "menuStopAutoRefresh";
            this.menuStopAutoRefresh.ShowImage = true;
            this.menuStopAutoRefresh.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.menuStopAutoRefresh_Click);
            //
            // Ribbon1
            //
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnToggleTaskPane;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton btnAutoRefresh;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTempo;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuAutoRefresh;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton menuRefreshNow;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton menuStopAutoRefresh;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
