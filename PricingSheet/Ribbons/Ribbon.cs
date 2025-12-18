using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using PricingSheet.Readers;

namespace PricingSheet.Ribbons
{
    public partial class Ribbon
    {
        public static Ribbon RibbonInstance { get; private set; }
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            RibbonInstance = this;
            Globals.ThisWorkbook.Application.SheetActivate += Application_SheetActivate;
            UpdateRibbonVisibility();
        }
        private void Application_SheetActivate(object Sh)
        {
            UpdateRibbonVisibility();
        }

        private void UpdateRibbonVisibility()
        {
            var activeSheet = Globals.ThisWorkbook.Application.ActiveSheet as ExcelInterop.Worksheet;
            this.tab1.Visible = true;
        }

        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            CSVReader reader = new CSVReader(Constants.PricingSheetFolderPath);
            Task.Run(() => MtM.MtMInstance.LoadAndDisplay(reader));
        }

        public Task SetStatus(string dbStatus = "", string spotStatus = "", string bbgStatus = "")
        {
            if (!string.IsNullOrEmpty(dbStatus))
                DbStatus.Label = dbStatus;
            if (!string.IsNullOrEmpty(spotStatus))
                SpotStatus.Label = spotStatus;
            if (!string.IsNullOrEmpty(bbgStatus))
                BbgConnection.Label = bbgStatus;

            return Task.CompletedTask;
        }

        public void SetAcrtiveSubscription(int count)
        {
            ActiveSubs.Label = $"Active Subscriptions: {count}/{Constants.MaxActiveInstruments}";
        }

        private void button5_Click_1(object sender, RibbonControlEventArgs e)
        {
            using (Forms.ManageSubs manageSubs = new Forms.ManageSubs())
            {
                manageSubs.ShowDialog();
            }

        }

        private void button6_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void EmailAlerts_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void Speak_Click(object sender, RibbonControlEventArgs e)
        {
        }

        private void toggleButton1_Click_1(object sender, RibbonControlEventArgs e)
        {

        }
    }
}
