namespace ExcelJsonTable
{
    partial class RibbonMenu : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonMenu()
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
            this.JsonTableRibbon = this.Factory.CreateRibbonTab();
            this.BasicGroup = this.Factory.CreateRibbonGroup();
            this.buttonImport = this.Factory.CreateRibbonButton();
            this.buttonExport = this.Factory.CreateRibbonButton();
            this.buttonAbout = this.Factory.CreateRibbonButton();
            this.buttonCreate = this.Factory.CreateRibbonButton();
            this.JsonTableRibbon.SuspendLayout();
            this.BasicGroup.SuspendLayout();
            //
            // JsonTableRibbon
            //
            this.JsonTableRibbon.Groups.Add(this.BasicGroup);
            this.JsonTableRibbon.Label = "JsonTable";
            this.JsonTableRibbon.Name = "JsonTableRibbon";
            //
            // BasicGroup
            //
            this.BasicGroup.Items.Add(this.buttonCreate);
            this.BasicGroup.Items.Add(this.buttonImport);
            this.BasicGroup.Items.Add(this.buttonAbout);
            this.BasicGroup.Items.Add(this.buttonExport);
            this.BasicGroup.Label = "Table";
            this.BasicGroup.Name = "BasicGroup";
            //
            // buttonImport
            //
            this.buttonImport.Label = "Import";
            this.buttonImport.Name = "buttonImport";
            this.buttonImport.OfficeImageId = "ImportTextFile";
            this.buttonImport.ShowImage = true;
            this.buttonImport.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonImport_Click);
            //
            // buttonExport
            //
            this.buttonExport.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonExport.Label = "Export";
            this.buttonExport.Name = "buttonExport";
            this.buttonExport.OfficeImageId = "ExportTextFile";
            this.buttonExport.ShowImage = true;
            this.buttonExport.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonExport_Click);
            //
            // buttonAbout
            //
            this.buttonAbout.Label = "About";
            this.buttonAbout.Name = "buttonAbout";
            this.buttonAbout.OfficeImageId = "HappyFace";
            this.buttonAbout.ShowImage = true;
            this.buttonAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonAbout_Click);
            //
            // buttonCreate
            //
            this.buttonCreate.Label = "Create";
            this.buttonCreate.Name = "buttonCreate";
            this.buttonCreate.OfficeImageId = "ImportTextFile";
            this.buttonCreate.ShowImage = true;
            this.buttonCreate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCreate_Click);
            //
            // RibbonMenu
            //
            this.Name = "RibbonMenu";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.JsonTableRibbon);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonMenu_Load);
            this.JsonTableRibbon.ResumeLayout(false);
            this.JsonTableRibbon.PerformLayout();
            this.BasicGroup.ResumeLayout(false);
            this.BasicGroup.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab JsonTableRibbon;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup BasicGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonExport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonImport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCreate;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonMenu RibbonMenu
        {
            get { return this.GetRibbon<RibbonMenu>(); }
        }
    }
}
