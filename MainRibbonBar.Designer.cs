
namespace VBA_Export
{
    partial class MainRibbonBar : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MainRibbonBar()
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
            this.VBAExport = this.Factory.CreateRibbonGroup();
            this.ExportButton = this.Factory.CreateRibbonButton();
            this.OpenFolderButton = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.VBAExport.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.VBAExport);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // VBAExport
            // 
            this.VBAExport.Items.Add(this.ExportButton);
            this.VBAExport.Items.Add(this.OpenFolderButton);
            this.VBAExport.Label = "VBA Export";
            this.VBAExport.Name = "VBAExport";
            // 
            // ExportButton
            // 
            this.ExportButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ExportButton.Label = "Export";
            this.ExportButton.Name = "ExportButton";
            this.ExportButton.OfficeImageId = "AfterInsert";
            this.ExportButton.ShowImage = true;
            this.ExportButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ExportButton_Click);
            // 
            // OpenFolderButton
            // 
            this.OpenFolderButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.OpenFolderButton.Label = "Open Folder";
            this.OpenFolderButton.Name = "OpenFolderButton";
            this.OpenFolderButton.OfficeImageId = "MasterExplorer";
            this.OpenFolderButton.ShowImage = true;
            this.OpenFolderButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OpenFolderButton_Click);
            // 
            // MainRibbonBar
            // 
            this.Name = "MainRibbonBar";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.VBAExport.ResumeLayout(false);
            this.VBAExport.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup VBAExport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ExportButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton OpenFolderButton;
    }

    partial class ThisRibbonCollection
    {
        internal MainRibbonBar Ribbon1
        {
            get { return this.GetRibbon<MainRibbonBar>(); }
        }
    }
}
