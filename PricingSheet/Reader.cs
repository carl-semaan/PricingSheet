using ClosedXML.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace PricingSheet
{
    public class Reader
    {
        public string FilePath { get; set; }
        public string FileName { get; set; }

        public Reader() { }

        public Reader(string filePath, string fileName)
        {
            this.FilePath = filePath;
            this.FileName = fileName;
        }
    }

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

    public class ExcelReader : Reader
    {
        public bool ReadHeaders { get; set; }
        public List<string> Headers { get; set; }

        public ExcelReader() { }
        public ExcelReader(string filePath, string fileName, bool ReadHeaders = true) : base(filePath, fileName)
        {
            this.ReadHeaders = true;
            this.Headers = new List<string>();
        }

        public List<T> LoadClass<T>() where T : new()
        {
            string fullPath = Path.Combine(FilePath, FileName);
            if (!File.Exists(fullPath))
                throw new FileNotFoundException(fullPath);

            var result = new List<T>();

            using (var workbook = new XLWorkbook(fullPath))
            {
                var ws = workbook.Worksheets.First();

                int startRow = 1;
                Dictionary<string, int> columnMap = new Dictionary<string, int>();

                if (ReadHeaders)
                {
                    var headerRow = ws.Row(1);
                    for (int col = 1; col <= ws.ColumnsUsed().Count(); col++)
                    {
                        string header = headerRow.Cell(col).GetValue<string>();
                        Headers.Add(header);
                        columnMap[header] = col;
                    }
                    startRow = 2;
                }

                var props = typeof(T).GetProperties();

                for (int row = startRow; row <= ws.RowsUsed().Count(); row++)
                {
                    T item = new T();
                    foreach (var prop in props)
                    {
                        string headerName = prop.Name;
                        if (columnMap.ContainsKey(headerName))
                        {
                            var cell = ws.Row(row).Cell(columnMap[headerName]);
                            try
                            {
                                if (prop.PropertyType == typeof(string))
                                {
                                    prop.SetValue(item, cell.GetValue<string>());
                                }
                                else if (prop.PropertyType == typeof(int) && int.TryParse(cell.GetValue<string>(), out int intVal))
                                {
                                    prop.SetValue(item, intVal);
                                }
                                else if (prop.PropertyType == typeof(double) && double.TryParse(cell.GetValue<string>(), out double dblVal))
                                {
                                    prop.SetValue(item, dblVal);
                                }
                                else if (prop.PropertyType == typeof(DateTime) && DateTime.TryParse(cell.GetValue<string>(), out DateTime dtVal))
                                {
                                    prop.SetValue(item, dtVal);
                                }
                            }
                            catch
                            {

                            }
                        }
                    }
                    result.Add(item);
                }
            }
            Headers = new List<string>();
            return result;
        }
    }

    public class CSVReader : Reader
    {
        public CSVReader() { }

        public CSVReader(string filePath, string fileName = "")
        {
            this.FilePath = filePath;
            this.FileName = fileName;
        }

        public CSVTicker LoadTickerData(string Ticker)
        {
            int maturityColStart = 6;

            if (string.IsNullOrEmpty(Ticker))
                throw new ArgumentException("Ticker cannot be null or empty.");

            string fullPath = Path.Combine(FilePath, $"{Ticker.ToUpper()}.csv");
            if (!File.Exists(fullPath))
                throw new FileNotFoundException(fullPath);

            var lines = File.ReadAllLines(fullPath);
            if (lines.Length < 2)
                throw new Exception("CSV file is empty or does not contain enough data.");

            CSVTicker tickerData = new CSVTicker();

            var headers = lines[0].Split(',').Skip(maturityColStart).ToList();
            var lastRow = lines.Last().Split(',').ToList();

            tickerData.Ticker = Ticker;
            tickerData.Date = lastRow[3];

            Dictionary<string, double> MaturityValues = new Dictionary<string, double>();
            for (int i = 0; i < headers.Count(); i++)
            {
                double value;
                try
                {
                    value = double.Parse(lastRow.ElementAt(i + maturityColStart));
                }
                catch
                {
                    value = double.NaN;
                }

                MaturityValues[headers[i]] = value;
            }

            tickerData.Maturities = MaturityValues;

            return tickerData;
        }
    }

    public class CSVTicker
    {
        public string Ticker { get; set; }
        public string Date { get; set; }
        public Dictionary<string, double> Maturities { get; set; }

        public CSVTicker()
        {
            Maturities = new Dictionary<string, double>();
        }

        public CSVTicker(string ticker, string ul, string currency, string date)
        {
            Ticker = ticker;
            Date = date;
            Maturities = new Dictionary<string, double>();
        }
    }
}
