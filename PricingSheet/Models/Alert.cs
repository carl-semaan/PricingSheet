using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PricingSheet.Models
{
    public class Alert
    {
        public string Instrument { get; set; }
        public string Underlying { get; set; }
        public string Maturity { get; set; }
        public string Field { get; set; }
        public AlertCondition Condition { get; set; }

        public Alert() { }

        public Alert(string instrument, string underlying, string maturity, string field, AlertCondition condition)
        {
            Instrument = instrument;
            Underlying = underlying;
            Maturity = maturity;
            Field = field;
            Condition = condition;
        }

        public enum AlertCondition
        {
            GreaterThan,
            LessThan
        }
    }
}
