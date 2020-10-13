namespace JdeScanExcelAddIn
{
    partial class JdeScanRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public JdeScanRibbon()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(JdeScanRibbon));
            this.tabJdeScan = this.Factory.CreateRibbonTab();
            this.grpJdeScan = this.Factory.CreateRibbonGroup();
            this.btnJdeScanExport = this.Factory.CreateRibbonButton();
            this.btnPlacePriority = this.Factory.CreateRibbonButton();
            this.tabJdeScan.SuspendLayout();
            this.grpJdeScan.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabJdeScan
            // 
            this.tabJdeScan.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabJdeScan.Groups.Add(this.grpJdeScan);
            this.tabJdeScan.Label = "JDE Scan";
            this.tabJdeScan.Name = "tabJdeScan";
            // 
            // grpJdeScan
            // 
            this.grpJdeScan.Items.Add(this.btnJdeScanExport);
            this.grpJdeScan.Items.Add(this.btnPlacePriority);
            this.grpJdeScan.Label = "JDE Scan";
            this.grpJdeScan.Name = "grpJdeScan";
            // 
            // btnJdeScanExport
            // 
            this.btnJdeScanExport.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnJdeScanExport.Image = ((System.Drawing.Image)(resources.GetObject("btnJdeScanExport.Image")));
            this.btnJdeScanExport.Label = "Export";
            this.btnJdeScanExport.Name = "btnJdeScanExport";
            this.btnJdeScanExport.ShowImage = true;
            this.btnJdeScanExport.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnJdeScanExport_Click);
            // 
            // btnPlacePriority
            // 
            this.btnPlacePriority.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnPlacePriority.Image = ((System.Drawing.Image)(resources.GetObject("btnPlacePriority.Image")));
            this.btnPlacePriority.Label = "Aktualizuj ABC";
            this.btnPlacePriority.Name = "btnPlacePriority";
            this.btnPlacePriority.ShowImage = true;
            this.btnPlacePriority.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPlacePriority_Click);
            // 
            // JdeScanRibbon
            // 
            this.Name = "JdeScanRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabJdeScan);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.JdeScanRibbon_Load);
            this.tabJdeScan.ResumeLayout(false);
            this.tabJdeScan.PerformLayout();
            this.grpJdeScan.ResumeLayout(false);
            this.grpJdeScan.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabJdeScan;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpJdeScan;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnJdeScanExport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPlacePriority;
    }

    partial class ThisRibbonCollection
    {
        internal JdeScanRibbon JdeScanRibbon
        {
            get { return this.GetRibbon<JdeScanRibbon>(); }
        }
    }
}
