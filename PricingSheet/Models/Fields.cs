using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PricingSheet.Models
{
    public class Fields
    {
        public string Field { get; set; }

        public Fields() { }

        public Fields(string Field)
        {
            this.Field = Field;
        }
    }
}
