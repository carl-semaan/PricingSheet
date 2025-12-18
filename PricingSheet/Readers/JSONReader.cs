using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PricingSheet.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PricingSheet.Readers
{
    public class JSONReader : Reader
    {
        public JSONReader() { }

        public JSONReader(string filePath, string fileName) : base(filePath, fileName) { }

        public List<T> LoadClass<T>(string targetName) where T : new()
        {
            string fullPath = Path.Combine(FilePath, FileName);
            if (!File.Exists(fullPath))
                throw new FileNotFoundException(fullPath);

            string jsonContent = File.ReadAllText(fullPath);
            var jObject = JObject.Parse(jsonContent);
            var jArray = jObject[targetName] as JArray;

            if (jArray == null)
                return new List<T>();

            return jArray.ToObject<List<T>>();
        }

        public void SaveJSON<T>(T data) where T : new()
        {
            string fullPath = Path.Combine(FilePath, FileName);
            if (!File.Exists(fullPath))
                throw new FileNotFoundException(fullPath);

            string jsonContent = JsonConvert.SerializeObject(data, Formatting.Indented);
            File.WriteAllText(fullPath, jsonContent);
        }
    }

    public class JSONContent
    {
        public List<Instruments> Instruments { get; set; }
        public List<Maturities> Maturities { get; set; }
        public List<Fields> Fields { get; set; }
        public List<LastPriceLoad> LastPriceLoad { get; set; }
        public List<UnderlyingSpot> UnderlyingSpot { get; set; }
        public JSONContent()
        {
            Instruments = new List<Instruments>();
            Maturities = new List<Maturities>();
            Fields = new List<Fields>();
            LastPriceLoad = new List<LastPriceLoad>();
            UnderlyingSpot = new List<UnderlyingSpot>();
        }

        public JSONContent(List<Instruments> instruments, List<Maturities> maturities, List<Fields> fields, List<LastPriceLoad> lastPriceLoad, List<UnderlyingSpot> underlyingSpot)
        {
            Instruments = instruments;
            Maturities = maturities;
            Fields = fields;
            LastPriceLoad = lastPriceLoad;
            UnderlyingSpot = underlyingSpot;
        }
    }
}
