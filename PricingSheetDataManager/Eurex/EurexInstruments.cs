using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CsvHelper.Configuration.Attributes;

namespace PricingSheetDataManager.Eurex
{
    public class EurexInstruments
    {
        public string ProductID { get; set; }
        public string Product { get; set; }

        [Name("Product Group")]
        public string ProductGroup {  get; set; }
        public string Currency { get; set; }

        [Name("Product ISIN")]
        public string ProductISIN { get; set; }

        [Name("Underlying ISIN")]
        public string UnderlyingISIN { get; set; }

        [Name("Share ISIN")]
        public string ShareISIN { get; set; }


    }
}
