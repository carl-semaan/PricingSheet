using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PricingSheet.Models
{
    public class UnderlyingSpot
    {
        public string Underlying { get; set; }
        public double? Value { get; set; }
        public UnderlyingSpot() { }
        public UnderlyingSpot(string Underlying, double? Value)
        {
            this.Underlying = Underlying;
            this.Value = Value;
        }
    }
}
