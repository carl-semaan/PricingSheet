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
    public partial class AddInstrumentForm : Form
    {
        public string TickerName => textBox1.Text;
        public string Underlying => textBox2.Text;
        public string ShortName => textBox3.Text;
        public string ExchangeCode => textBox4.Text;
        public string CurrencyISO => textBox5.Text;
        private readonly Flux _flux;
        public AddInstrumentForm(Flux fluxInstance)
        {
            _flux = fluxInstance;
            InitializeComponent();
        }
        public AddInstrumentForm()
        {
            InitializeComponent();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click_1(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
