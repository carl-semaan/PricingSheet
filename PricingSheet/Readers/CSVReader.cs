using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PricingSheet.Readers
{
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
