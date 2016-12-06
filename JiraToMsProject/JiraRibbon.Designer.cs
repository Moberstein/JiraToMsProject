namespace JiraToMsProject
{
    partial class JiraRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public JiraRibbon()
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
            this.group_jira = this.Factory.CreateRibbonGroup();
            this.button_import_jira = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group_jira.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group_jira);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group_jira
            // 
            this.group_jira.Items.Add(this.button_import_jira);
            this.group_jira.Label = "Jira";
            this.group_jira.Name = "group_jira";
            // 
            // button_import_jira
            // 
            this.button_import_jira.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_import_jira.Description = "Import Jira Tasks as XSL";
            this.button_import_jira.Image = global::JiraToMsProject.Properties.Resources.jira_logo;
            this.button_import_jira.Label = "Import Jira XSL";
            this.button_import_jira.Name = "button_import_jira";
            this.button_import_jira.ShowImage = true;
            // 
            // JiraRibbon
            // 
            this.Name = "JiraRibbon";
            this.RibbonType = "Microsoft.Project.Project";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.JiraRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group_jira.ResumeLayout(false);
            this.group_jira.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_jira;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_import_jira;
    }

    partial class ThisRibbonCollection
    {
        internal JiraRibbon JiraRibbon
        {
            get { return this.GetRibbon<JiraRibbon>(); }
        }
    }
}
