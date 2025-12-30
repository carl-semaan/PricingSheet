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
        [Name("PRODUCT_ID")]
        public string ProductID { get; set; }

        [Name("PRODUCT_NAME")]
        public string Product { get; set; }

        [Name("PRODUCT_GROUP")]
        public string ProductGroup {  get; set; }

        [Name("CURRENCY")]
        public string Currency { get; set; }

        [Name("PRODUCT_ISIN")]
        public string ProductISIN { get; set; }

        [Name("UNDERLYING_ISIN")]
        public string UnderlyingISIN { get; set; }

        [Name("SHARE_ISIN")]
        public string ShareISIN { get; set; }
    }
}
