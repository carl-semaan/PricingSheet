using DocumentFormat.OpenXml.Bibliography;
using PricingSheet.Models;
using PricingSheet.Readers;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PricingSheet.Forms
{
    public partial class EditMtM : Form
    {
        public List<Instruments> Instruments { get; set; }
        public List<Maturities> Maturities { get; set; }
        public List<CSVTicker> CSVTickers { get; set; }

        private List<CSVTicker> OriginalCopy;
        private List<CSVTicker> EditedTickers = new List<CSVTicker>();
        private DataTable Table;
        private BindingSource BindingSource;

        public EditMtM(List<Instruments> instruments, List<Maturities> maturities, List<CSVTicker> csvTickers)
        {
            InitializeComponent();
            this.ActiveControl = this.dataGridView1;

            Instruments = instruments;
            Maturities = maturities;
            CSVTickers = csvTickers.OrderBy(x => x.Ticker).ToList();
            OriginalCopy = CSVTickers.Select(x => x.Clone()).ToList();

            Task.Run(() =>
            {
                Table = GetDataTable();
                BindingSource = new BindingSource();
                BindingSource.DataSource = Table;

                dataGridView1.Invoke(new Action(() =>
                {
                    dataGridView1.DataSource = BindingSource;
                    dataGridView1.Columns["Ticker"].ReadOnly = true;
                }));
            });

        }

        public void GotFocus(object sender, EventArgs e)
        {
            if (SearchBox.Text == "Search...")
            {
                SearchBox.Text = "";
                SearchBox.ForeColor = Color.Black;
            }
        }

        public void LostFocus(object sender, EventArgs e)
        {
            if (SearchBox.Text == "")
            {
                SearchBox.Text = "Search...";
                SearchBox.ForeColor = Color.Gray;
            }
        }

        private void Save_Click(object sender, EventArgs e)
        {
            MtM.MtMInstance.RefreshSheet(CSVTickers);
            Univ.UnivInstance.UpdateFairValues(EditedTickers);
            EditedTickers.Clear();
            this.Close();
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            EditedTickers.Clear();
            CSVTickers.Clear();
            CSVTickers.AddRange(OriginalCopy.Select(x => x.Clone()));
            this.Close();
        }

        private DataTable GetDataTable()
        {
            DataTable table = new DataTable();
            table.BeginLoadData();

            table.Columns.Add("Ticker");

            var keys = CSVTickers
                .SelectMany(x => x.Maturities.Keys)
                .Where(x => !x.StartsWith("M"))
                .Distinct()
                .ToList();

            foreach (var key in keys)
                table.Columns.Add(key);

            foreach (var ticker in CSVTickers)
            {
                var row = table.NewRow();
                row["Ticker"] = ticker.Ticker;

                foreach (var key in keys)
                {
                    if (ticker.Maturities.TryGetValue(key, out double val))
                        row[key] = val;
                }

                table.Rows.Add(row);
            }

            table.EndLoadData();
            return table;
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            var row = dataGridView1.Rows[e.RowIndex];
            string tickerName = row.Cells["Ticker"].Value.ToString();
            string columnName = dataGridView1.Columns[e.ColumnIndex].Name;

            double value = row.Cells[e.ColumnIndex].Value != null ?
                           Convert.ToDouble(row.Cells[e.ColumnIndex].Value) : 0;

            // Find the original CSVTicker
            var original = CSVTickers.First(t => t.Ticker == tickerName);

            // Update the dictionary
            original.Maturities[columnName] = value;

            // Add to EditedTickers if not already added
            if (!EditedTickers.Any(t => t.Ticker == tickerName))
                EditedTickers.Add(original);
        }

        private void SearchBox_TextChanged(object sender, EventArgs e)
        {
            if (BindingSource == null) return;

            string filterText = SearchBox.Text.Replace("'", "''");

            if (string.IsNullOrWhiteSpace(filterText) || filterText == "Search...")
                BindingSource.RemoveFilter();
            else
                BindingSource.Filter = $"Ticker LIKE '%{filterText}%'";
        }
    }
}
