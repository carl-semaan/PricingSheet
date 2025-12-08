using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace PricingSheet.Ribbons
{
    public partial class FluxRibbon
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
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
            this.tab1.Visible = activeSheet?.Name == "Flux";
        }

        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {

        }
    }
}
