using ClosedXML.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
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
            this.ReadHeaders = ReadHeaders;
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
        /// <summary>
        /// High performance loading of multiple tickers from CSV files in parallel with least resource usage
        /// </summary>
        /// <param name="tickers"></param>
        /// <returns></returns>
        public async Task<List<CSVTicker>> LoadAllTickersAsync(IEnumerable<string> tickers)
        {
            var results = new ConcurrentBag<CSVTicker>();
            var missing = new ConcurrentBag<string>();

            int maxThreads = Math.Min(4, Environment.ProcessorCount);
            await Task.Run(() =>
            {
                Parallel.ForEach(tickers,
                    new ParallelOptions { MaxDegreeOfParallelism = maxThreads },
                    ticker =>
                    {
                        try
                        {
                            var data = LoadTickerData(ticker);
                            if (data != null)
                                results.Add(data);
                            else
                                missing.Add(ticker);
                        }
                        catch
                        {
                            missing.Add(ticker);
                        }
                    });
            });

            if (missing.Count > 0)
                Debug.WriteLine($"Missing tickers: {string.Join(", ", missing)}");

            return new List<CSVTicker>(results);
        }

        public CSVTicker LoadTickerData(string ticker)
        {
            string fullPath = Path.Combine(FilePath, $"{ticker.ToUpper()}.csv");

            if (!File.Exists(fullPath))
                throw new FileNotFoundException(fullPath);

            using (var sr = new StreamReader(fullPath))
            {
                string headerLine = sr.ReadLine();
                if (string.IsNullOrWhiteSpace(headerLine))
                    throw new Exception("CSV header missing");

                string[] headerParts = headerLine.Split(',');
                int maturityColStart = 6;
                int maturityCount = headerParts.Length - maturityColStart;

                string[] maturityLabels = new string[maturityCount];
                Array.Copy(headerParts, maturityColStart, maturityLabels, 0, maturityCount);

                string lastLine = null;
                while (!sr.EndOfStream)
                    lastLine = sr.ReadLine();

                if (string.IsNullOrWhiteSpace(lastLine))
                    throw new Exception("CSV data missing");

                string[] fields = lastLine.Split(',');
                CSVTicker tickerData = new CSVTicker
                {
                    Ticker = ticker,
                    Date = fields[3]
                };

                for (int i = 0; i < maturityLabels.Length; i++)
                {
                    string valStr = fields.Length > i + maturityColStart ? fields[i + maturityColStart] : "-";
                    if (double.TryParse(valStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double value))
                        tickerData.Maturities[maturityLabels[i]] = value;
                    else
                        tickerData.Maturities[maturityLabels[i]] = double.NaN;
                }

                return tickerData;
            }
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

        public CSVTicker(string ticker, string date)
        {
            Ticker = ticker;
            Date = date;
            Maturities = new Dictionary<string, double>();
        }
    }
}
