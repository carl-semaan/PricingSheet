using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PricingSheet.Models
{
    public class Maturities
    {
        public int Maturity { get; set; }
        public string MaturityCode { get; set; }
        public bool Flux { get; set; }
        public bool Active { get; set; }

        public Maturities() { }

        public Maturities(int Maturity, string MaturityCode, bool flux, bool active)
        {
            this.Maturity = Maturity;
            this.MaturityCode = MaturityCode;
            Flux = flux;
            Active = active;
        }
    }
}
