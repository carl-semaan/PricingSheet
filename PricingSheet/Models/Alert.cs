using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PricingSheet.Models
{
    public sealed class Alert : IEquatable<Alert>
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

        public bool Equals(Alert other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;

            return this.Instrument == other.Instrument &&
                   this.Underlying == other.Underlying &&
                   this.Maturity == other.Maturity &&
                   this.Field == other.Field &&
                   this.Condition == other.Condition;
        }

        public override bool Equals(object obj) => Equals(obj as Alert);

        public override int GetHashCode()
        {
            unchecked
            {
                int hash = 17;
                hash = hash * 23 + (Instrument?.GetHashCode() ?? 0);
                hash = hash * 23 + (Underlying?.GetHashCode() ?? 0);
                hash = hash * 23 + (Maturity?.GetHashCode() ?? 0);
                hash = hash * 23 + (Field?.GetHashCode() ?? 0);
                hash = hash * 23 + Condition.GetHashCode();
                return hash;
            }
        }

        public enum AlertCondition
        {
            GreaterThan,
            LessThan
        }
    }
}
