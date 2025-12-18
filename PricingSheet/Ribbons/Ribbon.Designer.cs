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
            this.group1 = this.Factory.CreateRibbonGroup();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.Status = this.Factory.CreateRibbonGroup();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.label2 = this.Factory.CreateRibbonLabel();
            this.label3 = this.Factory.CreateRibbonLabel();
            this.DbStatus = this.Factory.CreateRibbonLabel();
            this.SpotStatus = this.Factory.CreateRibbonLabel();
            this.BbgConnection = this.Factory.CreateRibbonLabel();
            this.SubManager = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.button4 = this.Factory.CreateRibbonButton();
            this.Refresh = this.Factory.CreateRibbonButton();
            this.ActiveSubs = this.Factory.CreateRibbonButton();
            this.button5 = this.Factory.CreateRibbonButton();
            this.Notifications = this.Factory.CreateRibbonGroup();
            this.EmailAlerts = this.Factory.CreateRibbonToggleButton();
            this.SpeechAlerts = this.Factory.CreateRibbonToggleButton();
            this.button6 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.Status.SuspendLayout();
            this.SubManager.SuspendLayout();
            this.Notifications.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.Status);
            this.tab1.Groups.Add(this.SubManager);
            this.tab1.Groups.Add(this.Notifications);
            this.tab1.Label = "Pricing Sheet";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button1);
            this.group1.Items.Add(this.button2);
            this.group1.Label = "Instruments";
            this.group1.Name = "group1";
            // 
            // group2
            // 
            this.group2.Items.Add(this.button3);
            this.group2.Items.Add(this.button4);
            this.group2.Label = "Maturities";
            this.group2.Name = "group2";
            // 
            // group3
            // 
            this.group3.Items.Add(this.Refresh);
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
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.Label = "Add";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
            this.button2.Label = "Edit";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            // 
            // button3
            // 
            this.button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button3.Image = ((System.Drawing.Image)(resources.GetObject("button3.Image")));
            this.button3.Label = "Add";
            this.button3.Name = "button3";
            this.button3.ShowImage = true;
            // 
            // button4
            // 
            this.button4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button4.Image = ((System.Drawing.Image)(resources.GetObject("button4.Image")));
            this.button4.Label = "Edit";
            this.button4.Name = "button4";
            this.button4.ShowImage = true;
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
            // Notifications
            // 
            this.Notifications.Items.Add(this.EmailAlerts);
            this.Notifications.Items.Add(this.SpeechAlerts);
            this.Notifications.Items.Add(this.button6);
            this.Notifications.Label = "Notifications";
            this.Notifications.Name = "Notifications";
            // 
            // EmailAlerts
            // 
            this.EmailAlerts.Checked = true;
            this.EmailAlerts.Label = "Email Alerts";
            this.EmailAlerts.Name = "EmailAlerts";
            this.EmailAlerts.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.EmailAlerts_Click);
            // 
            // SpeechAlerts
            // 
            this.SpeechAlerts.Checked = true;
            this.SpeechAlerts.Label = "Speech Alerts";
            this.SpeechAlerts.Name = "SpeechAlerts";
            this.SpeechAlerts.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButton1_Click_1);
            // 
            // button6
            // 
            this.button6.Label = "Speech test";
            this.button6.Name = "button6";
            this.button6.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button6_Click_1);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
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
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal RibbonGroup group2;
        internal RibbonButton button3;
        internal RibbonButton button4;
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
        internal RibbonToggleButton EmailAlerts;
        internal RibbonToggleButton SpeechAlerts;
        internal RibbonButton button6;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon FluxRibbon
        {
            get { return Globals.Factory.GetRibbonFactory().CreateRibbonTab() as Ribbon; }

        }

    }
}
