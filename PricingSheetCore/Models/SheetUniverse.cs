using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PricingSheetCore.Models
{
    public class SheetUniverse
    {
        public List<Instruments> Instruments { get; set; }
        public List<Maturities> Maturities { get; set; }
        public List<Fields> Fields { get; set; }

        public SheetUniverse() { }

        public SheetUniverse(List<Instruments> instruments, List<Maturities> maturities)
        {
            Instruments = instruments;
            Maturities = maturities;
        }

        public SheetUniverse(List<Instruments> instruments, List<Maturities> maturities, List<Fields> fields)
        {
            Instruments = instruments;
            Maturities = maturities;
            Fields = fields;
        }
    }
}
