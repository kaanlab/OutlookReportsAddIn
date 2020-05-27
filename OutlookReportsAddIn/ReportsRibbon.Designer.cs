namespace OutlookReportsAddIn
{
    partial class ReportsRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ReportsRibbon()
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
            this.reportsAddIn = this.Factory.CreateRibbonGroup();
            this.OneDayReport = this.Factory.CreateRibbonButton();
            this.ReportForPeriod = this.Factory.CreateRibbonButton();
            this.SettingsButton = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.reportsAddIn.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.reportsAddIn);
            this.tab1.Label = "Отчеты";
            this.tab1.Name = "tab1";
            // 
            // reportsAddIn
            // 
            this.reportsAddIn.Items.Add(this.OneDayReport);
            this.reportsAddIn.Items.Add(this.ReportForPeriod);
            this.reportsAddIn.Items.Add(this.SettingsButton);
            this.reportsAddIn.Label = "Сформировать отчет";
            this.reportsAddIn.Name = "reportsAddIn";
            // 
            // OneDayReport
            // 
            this.OneDayReport.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.OneDayReport.Label = "За день";
            this.OneDayReport.Name = "OneDayReport";
            this.OneDayReport.OfficeImageId = "ManageReplies";
            this.OneDayReport.ShowImage = true;
            this.OneDayReport.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OneDayReport_Click);
            // 
            // ReportForPeriod
            // 
            this.ReportForPeriod.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ReportForPeriod.Label = "За период";
            this.ReportForPeriod.Name = "ReportForPeriod";
            this.ReportForPeriod.OfficeImageId = "CreateEmail";
            this.ReportForPeriod.ShowImage = true;
            this.ReportForPeriod.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ReportForPeriod_Click);
            // 
            // SettingsButton
            // 
            this.SettingsButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.SettingsButton.Label = "Настройки";
            this.SettingsButton.Name = "SettingsButton";
            this.SettingsButton.OfficeImageId = "FileProperties";
            this.SettingsButton.ShowImage = true;
            this.SettingsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SettingsButton_Click);
            // 
            // ReportsRibbon
            // 
            this.Name = "ReportsRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ReportsRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.reportsAddIn.ResumeLayout(false);
            this.reportsAddIn.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup reportsAddIn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton OneDayReport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SettingsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ReportForPeriod;
    }

    partial class ThisRibbonCollection
    {
        internal ReportsRibbon ReportsRibbon
        {
            get { return this.GetRibbon<ReportsRibbon>(); }
        }
    }
}
