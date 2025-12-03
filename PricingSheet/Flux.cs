using Microsoft.VisualStudio.Tools.Applications.Runtime;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace PricingSheet
{
    public partial class Flux
    {
        private void Sheet3_Startup(object sender, System.EventArgs e)
        {
            SheetButton sheetButton = new SheetButton(
                "Say Hello",
                1,
                1,
                "Blue",
                () => System.Windows.Forms.MessageBox.Show("Welcome to the new Pricing Sheet!!!"));

            List<SheetButton> ButtonsList = new List<SheetButton>();
            ButtonsList.Add(sheetButton);

            var sheet = Globals.ThisWorkbook.Worksheets["Sheet3"];
            var vstoSheet = Globals.Factory.GetVstoObject(sheet);

            SheetInitialization sheetInitialization = new SheetInitialization(
                vstoSheet,
                "Flux",
                true,
                ButtonsList
            );

            sheetInitialization.Run();
        }

        private void Sheet3_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Sheet3_Startup);
            this.Shutdown += new System.EventHandler(Sheet3_Shutdown);
        }

        #endregion

    }
}
