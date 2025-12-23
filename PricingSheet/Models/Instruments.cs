using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PricingSheet.Models
{
    public class Instruments
    {
        public string Ticker { get; set; }
        public string Underlying { get; set; }
        public string ShortName { get; set; }
        public string ExchangeCode { get; set; }
        public string InstrumentType { get; set; }
        public string Currency { get; set; }
        public string ICBSuperSectorName { get; set; }

        public Instruments() { }
        public Instruments(string ticker, string underlying, string shortName, string exchangeCode, string instrumentType, string currency, string ICBSupersectorName)
        {
            Ticker = ticker;
            Underlying = underlying;
            ShortName = shortName;
            ExchangeCode = exchangeCode;
            InstrumentType = instrumentType;
            Currency = currency;
            this.ICBSuperSectorName = ICBSupersectorName;
        }
    }
}
