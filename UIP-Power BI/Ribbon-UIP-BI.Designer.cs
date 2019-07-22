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
            this.General_Group = this.Factory.CreateRibbonGroup();
            this.TradeUpdates_Group = this.Factory.CreateRibbonGroup();
            this.Add_New_Trade_Button = this.Factory.CreateRibbonButton();
            this.View_Settings_Button = this.Factory.CreateRibbonButton();
            this.View_Index_Button = this.Factory.CreateRibbonButton();
            this.Create_Backup_Button = this.Factory.CreateRibbonButton();
            this.CreateTradeExportTable_Button = this.Factory.CreateRibbonButton();
            this.UpdateTrade_Button = this.Factory.CreateRibbonButton();
            this.UpdateTradeSchedule_Button = this.Factory.CreateRibbonButton();
            this.UIP_BI.SuspendLayout();
            this.General_Group.SuspendLayout();
            this.TradeUpdates_Group.SuspendLayout();
            this.SuspendLayout();
            // 
            // UIP_BI
            // 
            this.UIP_BI.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.UIP_BI.Groups.Add(this.General_Group);
            this.UIP_BI.Groups.Add(this.TradeUpdates_Group);
            this.UIP_BI.Label = "UIP BI";
            this.UIP_BI.Name = "UIP_BI";
            // 
            // General_Group
            // 
            this.General_Group.Items.Add(this.Add_New_Trade_Button);
            this.General_Group.Items.Add(this.View_Settings_Button);
            this.General_Group.Items.Add(this.View_Index_Button);
            this.General_Group.Items.Add(this.Create_Backup_Button);
            this.General_Group.Label = "General";
            this.General_Group.Name = "General_Group";
            // 
            // TradeUpdates_Group
            // 
            this.TradeUpdates_Group.Items.Add(this.CreateTradeExportTable_Button);
            this.TradeUpdates_Group.Items.Add(this.UpdateTradeSchedule_Button);
            this.TradeUpdates_Group.Items.Add(this.UpdateTrade_Button);
            this.TradeUpdates_Group.Label = "Trade Updates";
            this.TradeUpdates_Group.Name = "TradeUpdates_Group";
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
            this.View_Settings_Button.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.View_Settings_Button_Click);
            // 
            // View_Index_Button
            // 
            this.View_Index_Button.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.View_Index_Button.Label = "View Index";
            this.View_Index_Button.Name = "View_Index_Button";
            this.View_Index_Button.OfficeImageId = "AccessRelinkLists";
            this.View_Index_Button.ShowImage = true;
            this.View_Index_Button.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.View_Index_Button_Click);
            // 
            // Create_Backup_Button
            // 
            this.Create_Backup_Button.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Create_Backup_Button.Label = "Create Backup Copy";
            this.Create_Backup_Button.Name = "Create_Backup_Button";
            this.Create_Backup_Button.OfficeImageId = "Archive";
            this.Create_Backup_Button.ShowImage = true;
            this.Create_Backup_Button.SuperTip = "Saves a dated copy of this file to a \"Backup Copies\" folder at the same location " +
    "as the current file.";
            // 
            // CreateTradeExportTable_Button
            // 
            this.CreateTradeExportTable_Button.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.CreateTradeExportTable_Button.Label = "Create Trade Export Table";
            this.CreateTradeExportTable_Button.Name = "CreateTradeExportTable_Button";
            this.CreateTradeExportTable_Button.OfficeImageId = "AdpDiagramAddTable";
            this.CreateTradeExportTable_Button.ShowImage = true;
            // 
            // UpdateTrade_Button
            // 
            this.UpdateTrade_Button.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.UpdateTrade_Button.Label = "Update Trade";
            this.UpdateTrade_Button.Name = "UpdateTrade_Button";
            this.UpdateTrade_Button.OfficeImageId = "AccessListTasks";
            this.UpdateTrade_Button.ShowImage = true;
            // 
            // UpdateTradeSchedule_Button
            // 
            this.UpdateTradeSchedule_Button.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.UpdateTradeSchedule_Button.Label = "Update Trade Schedule";
            this.UpdateTradeSchedule_Button.Name = "UpdateTradeSchedule_Button";
            this.UpdateTradeSchedule_Button.OfficeImageId = "CalendarViewGallery";
            this.UpdateTradeSchedule_Button.ShowImage = true;
            // 
            // Ribbon_UIP_BI
            // 
            this.Name = "Ribbon_UIP_BI";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.UIP_BI);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_UIP_BI_Load);
            this.UIP_BI.ResumeLayout(false);
            this.UIP_BI.PerformLayout();
            this.General_Group.ResumeLayout(false);
            this.General_Group.PerformLayout();
            this.TradeUpdates_Group.ResumeLayout(false);
            this.TradeUpdates_Group.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab UIP_BI;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup General_Group;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Add_New_Trade_Button;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton View_Settings_Button;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Create_Backup_Button;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup TradeUpdates_Group;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton View_Index_Button;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CreateTradeExportTable_Button;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton UpdateTrade_Button;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton UpdateTradeSchedule_Button;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon_UIP_BI Ribbon_UIP_BI
        {
            get { return this.GetRibbon<Ribbon_UIP_BI>(); }
        }
    }
}
