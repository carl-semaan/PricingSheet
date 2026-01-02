using Microsoft.Office.Tools.Ribbon;
using PricingSheet.Forms;

namespace PricingSheet.Ribbons
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.Status = this.Factory.CreateRibbonGroup();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.label2 = this.Factory.CreateRibbonLabel();
            this.label3 = this.Factory.CreateRibbonLabel();
            this.DbStatus = this.Factory.CreateRibbonLabel();
            this.SpotStatus = this.Factory.CreateRibbonLabel();
            this.BbgConnection = this.Factory.CreateRibbonLabel();
            this.SubManager = this.Factory.CreateRibbonGroup();
            this.Notifications = this.Factory.CreateRibbonGroup();
            this.Refresh = this.Factory.CreateRibbonButton();
            this.EditMtM = this.Factory.CreateRibbonButton();
            this.ActiveSubs = this.Factory.CreateRibbonButton();
            this.button5 = this.Factory.CreateRibbonButton();
            this.Alerts = this.Factory.CreateRibbonToggleButton();
            this.tab1.SuspendLayout();
            this.group3.SuspendLayout();
            this.Status.SuspendLayout();
            this.SubManager.SuspendLayout();
            this.Notifications.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.Status);
            this.tab1.Groups.Add(this.SubManager);
            this.tab1.Groups.Add(this.Notifications);
            this.tab1.Label = "Pricing Sheet";
            this.tab1.Name = "tab1";
            // 
            // group3
            // 
            this.group3.Items.Add(this.Refresh);
            this.group3.Items.Add(this.EditMtM);
            this.group3.Label = "MtM Sheet";
            this.group3.Name = "group3";
            // 
            // Status
            // 
            this.Status.Items.Add(this.label1);
            this.Status.Items.Add(this.label2);
            this.Status.Items.Add(this.label3);
            this.Status.Items.Add(this.DbStatus);
            this.Status.Items.Add(this.SpotStatus);
            this.Status.Items.Add(this.BbgConnection);
            this.Status.Label = "Sheet Status";
            this.Status.Name = "Status";
            // 
            // label1
            // 
            this.label1.Label = "Database Prices:";
            this.label1.Name = "label1";
            // 
            // label2
            // 
            this.label2.Label = "Spot Prices:";
            this.label2.Name = "label2";
            // 
            // label3
            // 
            this.label3.Label = "Bloomberg Pipeline:";
            this.label3.Name = "label3";
            // 
            // DbStatus
            // 
            this.DbStatus.Label = "Pending";
            this.DbStatus.Name = "DbStatus";
            // 
            // SpotStatus
            // 
            this.SpotStatus.Label = "Pending";
            this.SpotStatus.Name = "SpotStatus";
            // 
            // BbgConnection
            // 
            this.BbgConnection.Label = "Pending";
            this.BbgConnection.Name = "BbgConnection";
            // 
            // SubManager
            // 
            this.SubManager.Items.Add(this.ActiveSubs);
            this.SubManager.Items.Add(this.button5);
            this.SubManager.Label = "Subscription Manager";
            this.SubManager.Name = "SubManager";
            // 
            // Notifications
            // 
            this.Notifications.Items.Add(this.Alerts);
            this.Notifications.Label = "Notifications";
            this.Notifications.Name = "Notifications";
            // 
            // Refresh
            // 
            this.Refresh.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Refresh.Image = global::PricingSheet.Properties.Resources.refresh_page_option;
            this.Refresh.Label = "Refresh";
            this.Refresh.Name = "Refresh";
            this.Refresh.ShowImage = true;
            this.Refresh.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button5_Click);
            // 
            // EditMtM
            // 
            this.EditMtM.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.EditMtM.Image = ((System.Drawing.Image)(resources.GetObject("EditMtM.Image")));
            this.EditMtM.Label = "Edit";
            this.EditMtM.Name = "EditMtM";
            this.EditMtM.ShowImage = true;
            this.EditMtM.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button6_Click_2);
            // 
            // ActiveSubs
            // 
            this.ActiveSubs.Image = ((System.Drawing.Image)(resources.GetObject("ActiveSubs.Image")));
            this.ActiveSubs.Label = "Active Subscriptions: 0/0";
            this.ActiveSubs.Name = "ActiveSubs";
            this.ActiveSubs.ShowImage = true;
            this.ActiveSubs.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button6_Click);
            // 
            // button5
            // 
            this.button5.Image = ((System.Drawing.Image)(resources.GetObject("button5.Image")));
            this.button5.Label = "Manage Subscriptions";
            this.button5.Name = "button5";
            this.button5.ShowImage = true;
            this.button5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button5_Click_1);
            // 
            // Alerts
            // 
            this.Alerts.Checked = true;
            this.Alerts.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Alerts.Image = ((System.Drawing.Image)(resources.GetObject("Alerts.Image")));
            this.Alerts.Label = "Alerts";
            this.Alerts.Name = "Alerts";
            this.Alerts.ShowImage = true;
            this.Alerts.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ToggleSpeechAlert);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.Status.ResumeLayout(false);
            this.Status.PerformLayout();
            this.SubManager.ResumeLayout(false);
            this.SubManager.PerformLayout();
            this.Notifications.ResumeLayout(false);
            this.Notifications.PerformLayout();
            this.ResumeLayout(false);

        }

        private void AddInstrument(object sender, RibbonControlEventArgs e)
        {
            using(var form = new Forms.AddInstrumentForm())
            {
                form.ShowDialog();
            }
        }
        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal RibbonGroup group3;
        internal RibbonButton Refresh;
        internal RibbonGroup Status;
        internal RibbonLabel label1;
        internal RibbonLabel label2;
        internal RibbonLabel label3;
        internal RibbonLabel DbStatus;
        internal RibbonLabel SpotStatus;
        internal RibbonLabel BbgConnection;
        internal RibbonGroup SubManager;
        internal RibbonButton button5;
        internal RibbonButton ActiveSubs;
        internal RibbonGroup Notifications;
        internal RibbonToggleButton Alerts;
        internal RibbonButton EditMtM;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon FluxRibbon
        {
            get { return Globals.Factory.GetRibbonFactory().CreateRibbonTab() as Ribbon; }

        }

    }
}
