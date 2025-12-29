using PricingSheetCore.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using PricingSheetCore.Readers;
using PricingSheetCore;

namespace PricingSheet.Forms
{
    public partial class ManageSubs : Form
    {
        public List<Instruments> Instruments { get; set; }
        public List<Maturities> Maturities { get; set; }
        public ManageSubs()
        {
            InitializeComponent();

            JSONReader reader = new JSONReader(Constants.PricingSheetFolderPath, Constants.JSONFileName);
            Instruments = reader.LoadClass<Instruments>(nameof(Instruments));
            Maturities = reader.LoadClass<Maturities>(nameof(Maturities)).Where(x => x.Flux).ToList();

            this.checkedListBox1.Items.AddRange(Maturities.Select(x => x.MaturityCode).ToArray());

            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                var checkbox = checkedListBox1.Items[i];
                if (Maturities.First(x => x.MaturityCode == (string)checkbox).Active)
                    checkedListBox1.SetItemChecked(checkedListBox1.Items.IndexOf(checkbox), true);
            }

            DisplayStats();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < Maturities.Count; i++)
                checkedListBox1.SetItemChecked(i, true);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            for(int i = 0; i < Maturities.Count; i++)
                checkedListBox1.SetItemChecked(i, false);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Task.Run(() =>
            {
                // Saving the updated flags the the json file
                JSONReader reader = new JSONReader(Constants.PricingSheetFolderPath, Constants.JSONFileName);

                List<Maturities> newMaturities = reader.LoadClass<Maturities>(nameof(Maturities));
                newMaturities.ForEach(mat =>
                {
                    var updatedMat = Maturities.FirstOrDefault(x => x.MaturityCode == mat.MaturityCode && x.Maturity == mat.Maturity);
                    if(updatedMat != null)
                        mat.Active = updatedMat.Active;
                });

                JSONContent content = new JSONContent(
                    reader.LoadClass<Instruments>(nameof(Instruments)),
                    newMaturities,
                    reader.LoadClass<Fields>(nameof(Fields)),
                    reader.LoadClass<LastPriceLoad>(nameof(LastPriceLoad)),
                    reader.LoadClass<UnderlyingSpot>(nameof(UnderlyingSpot))
                    );

                reader.SaveJSON(content);

                // Updating the flux sheet data
                Flux.FluxInstance.UpdateSubscriptions(newMaturities);

                // Updating the univ sheet data
                Univ.UnivInstance.UpdateSubscriptions(newMaturities);
            });

            this.Close();
        }

        private void CheckedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            // Get the item that is being clicked
            var item = (string)checkedListBox1.Items[e.Index];

            // Determine the new check state
            bool willBeChecked = e.NewValue == CheckState.Checked;

            // Set the Active flag based on the new check state
            Maturities.First(x => x.MaturityCode == item).Active = willBeChecked;

            // Perform calculations based on this item
            DisplayStats();
        }

        private void DisplayStats()
        {
            int instrCount = Instruments.Count;
            int matCount = Maturities.Count(x => x.Active);
            int total = matCount * instrCount;

            this.Instr.Text = $"Instruments: {instrCount}";
            this.Mat.Text = $"Maturities: {matCount}";
            this.Subscriptions.Text = $"Subscriptions: {total} / {Constants.MaxActiveInstruments}";

            if (total > Constants.MaxActiveInstruments)
            {
                this.Subscriptions.ForeColor = System.Drawing.Color.Red;
                this.button3.Enabled = false;
            }
            else
            {
                this.Subscriptions.ForeColor = System.Drawing.Color.Green;
                this.button3.Enabled = true;
            }
        }
    }
}
