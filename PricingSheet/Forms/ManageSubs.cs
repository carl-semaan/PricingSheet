using PricingSheet.Models;
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
    public partial class ManageSubs : Form
    {
        public List<Instruments> Instruments { get; set; }
        public List<Maturities> Maturities { get; set; }
        public ManageSubs()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
