namespace UIP_Power_BI
{
    partial class Ribbon_UIP_BI : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon_UIP_BI()
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
            this.UIP_BI = this.Factory.CreateRibbonTab();
            this.Setup_Group = this.Factory.CreateRibbonGroup();
            this.Backup_Copy_SplitButton = this.Factory.CreateRibbonSplitButton();
            this.Add_New_Trade_Button = this.Factory.CreateRibbonButton();
            this.View_Settings_Button = this.Factory.CreateRibbonButton();
            this.Create_Backup_Button = this.Factory.CreateRibbonButton();
            this.UIP_BI.SuspendLayout();
            this.Setup_Group.SuspendLayout();
            this.SuspendLayout();
            // 
            // UIP_BI
            // 
            this.UIP_BI.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.UIP_BI.Groups.Add(this.Setup_Group);
            this.UIP_BI.Label = "UIP BI";
            this.UIP_BI.Name = "UIP_BI";
            // 
            // Setup_Group
            // 
            this.Setup_Group.Items.Add(this.Add_New_Trade_Button);
            this.Setup_Group.Items.Add(this.View_Settings_Button);
            this.Setup_Group.Items.Add(this.Backup_Copy_SplitButton);
            this.Setup_Group.Label = "Setup";
            this.Setup_Group.Name = "Setup_Group";
            // 
            // Backup_Copy_SplitButton
            // 
            this.Backup_Copy_SplitButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Backup_Copy_SplitButton.Items.Add(this.Create_Backup_Button);
            this.Backup_Copy_SplitButton.Label = "Create Backup Copy";
            this.Backup_Copy_SplitButton.Name = "Backup_Copy_SplitButton";
            this.Backup_Copy_SplitButton.OfficeImageId = "Archive";
            this.Backup_Copy_SplitButton.SuperTip = "Saves a dated copy of this file to a \"Backup Copies\" folder at the same location " +
    "as the current file.";
            // 
            // Add_New_Trade_Button
            // 
            this.Add_New_Trade_Button.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Add_New_Trade_Button.Label = "Add New Trade";
            this.Add_New_Trade_Button.Name = "Add_New_Trade_Button";
            this.Add_New_Trade_Button.OfficeImageId = "AddAccount";
            this.Add_New_Trade_Button.ShowImage = true;
            this.Add_New_Trade_Button.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Add_New_Trade_Button_Click);
            // 
            // View_Settings_Button
            // 
            this.View_Settings_Button.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.View_Settings_Button.Label = "View Settings";
            this.View_Settings_Button.Name = "View_Settings_Button";
            this.View_Settings_Button.OfficeImageId = "AddInManager";
            this.View_Settings_Button.ShowImage = true;
            // 
            // Create_Backup_Button
            // 
            this.Create_Backup_Button.Label = "Create Backup Copy";
            this.Create_Backup_Button.Name = "Create_Backup_Button";
            this.Create_Backup_Button.OfficeImageId = "Archive";
            this.Create_Backup_Button.ShowImage = true;
            this.Create_Backup_Button.SuperTip = "Saves a dated copy of this file to a \"Backup Copies\" folder at the same location " +
    "as the current file.";
            // 
            // Ribbon_UIP_BI
            // 
            this.Name = "Ribbon_UIP_BI";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.UIP_BI);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_UIP_BI_Load);
            this.UIP_BI.ResumeLayout(false);
            this.UIP_BI.PerformLayout();
            this.Setup_Group.ResumeLayout(false);
            this.Setup_Group.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab UIP_BI;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Setup_Group;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton Backup_Copy_SplitButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Add_New_Trade_Button;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton View_Settings_Button;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Create_Backup_Button;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon_UIP_BI Ribbon_UIP_BI
        {
            get { return this.GetRibbon<Ribbon_UIP_BI>(); }
        }
    }
}
