namespace PowerPointWSRemote
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
            this.txtPort = this.Factory.CreateRibbonEditBox();
            this.cbEnabled = this.Factory.CreateRibbonCheckBox();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "Remote";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.cbEnabled);
            this.group1.Items.Add(this.txtPort);
            this.group1.Label = "Settings";
            this.group1.Name = "group1";
            // 
            // txtPort
            // 
            this.txtPort.Label = "Port:";
            this.txtPort.Name = "txtPort";
            this.txtPort.Text = global::PowerPointWSRemote.Properties.Settings.Default.port;
            this.txtPort.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.txtPort_TextChanged);
            // 
            // cbEnabled
            // 
            this.cbEnabled.Label = "Enable";
            this.cbEnabled.Name = "cbEnabled";
            this.cbEnabled.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbEnabled_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
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
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox txtPort;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbEnabled;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
