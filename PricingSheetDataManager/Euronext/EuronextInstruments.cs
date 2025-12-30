using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CsvHelper.Configuration.Attributes;

namespace PricingSheetDataManager.Euronext
{
    public class EuronextInstruments
    {
        [Name("Instrument name")]
        public string InstrumentName { get; set; }

        [Name("Code")]
        public string Code { get; set; }

        [Name("Product family")]
        public string ProductFamily { get; set; }

        [Name("Underlying ISIN")]
        public string UnderlyingISIN { get; set; }

        [Name("Location")]
        public string Location { get; set; }
    }
}
