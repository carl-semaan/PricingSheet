using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PricingSheet.Models
{
    public class LastPriceLoad
    {
        public DateTime LastLoad { get; set; }
        public LastPriceLoad() { }
        public LastPriceLoad(DateTime LastLoad)
        {
            this.LastLoad = LastLoad;
        }
    }
}
