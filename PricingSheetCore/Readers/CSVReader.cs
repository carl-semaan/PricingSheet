using CsvHelper;
using CsvHelper.Configuration;
using PricingSheetCore.Models;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PricingSheetCore.Readers
{
    public class CSVReader : Reader
    {
        public string Delimiter { get; set; }
        public bool SkipFirstRow { get; set; }
        public CSVReader() { }

        public CSVReader(string filePath, string fileName = "", string Delimiter = ",", bool SkipFirstRow = false)
        {
            this.FilePath = filePath;
            this.FileName = fileName;
            this.Delimiter = Delimiter;
            this.SkipFirstRow = SkipFirstRow;
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

        public void SaveTickerData(List<CSVTicker> newValues, List<Instruments> MtMInstruments, Dictionary<string, UnderlyingSpot> MtMUnderlyingSpot)
        {
            foreach (var ticker in newValues)
            {
                try
                {
                    string fullPath = Path.Combine(FilePath, $"{ticker.Ticker.ToUpper()}.csv");

                    Instruments targetInstrument = MtMInstruments.First(x => x.Ticker == ticker.Ticker);
                    UnderlyingSpot targetSpot = MtMUnderlyingSpot.First(x => x.Key == targetInstrument.Underlying).Value;

                    // Getting the headers
                    var headerLine = File.ReadLines(fullPath).First();
                    string[] headers = headerLine.Split(',');

                    // Getting the values and mapping them to the headers in the file
                    var row = new Dictionary<string, string>
                    {
                        [headers[0]] = ticker.Ticker,
                        [headers[1]] = targetInstrument.GetUlRtCode(),
                        [headers[2]] = targetInstrument.Currency,
                        [headers[3]] = ticker.Date,
                        [headers[4]] = targetSpot.Value.ToString()
                    };

                    foreach (var h in headers.Skip(5))
                    {
                        row[h] = "-";

                        try
                        {
                            if (!double.IsNaN(ticker.Maturities[h]))
                                row[h] = ticker.Maturities[h].ToString();
                        }
                        catch { }
                    }

                    // Building the line to append and making sure to follow header's order
                    var csvLine = string.Join(",", headers.Select(h => row.TryGetValue(h, out var v) ? v : ""));

                    using (var sw = new StreamWriter(fullPath, append: true))
                    {
                        sw.WriteLine(csvLine);
                    }
                }
                catch { }
            }
        }

        public List<T> LoadClass<T>()
        {
            string fullPath = Path.Combine(FilePath, FileName);

            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HasHeaderRecord = true,
                Delimiter = this.Delimiter,
            };

            using var reader = new StreamReader(fullPath);
            using var csv = new CsvReader(reader, config);

            csv.Context.Configuration.HeaderValidated = null;
            csv.Context.Configuration.MissingFieldFound = null;

            if (SkipFirstRow)
            {
                csv.Read();
            }

            return csv.GetRecords<T>().ToList();
        }

        public void AddMaturities(List<Maturities> maturities)
        {
            List<string> maturityCodes = new List<string>();
            foreach (var mat in maturities)
            {
                string year = mat.MaturityCode.Substring(1, 2);
                maturityCodes.Add($"M{year}");
                maturityCodes.Add($"Z{year}");
            }

            int ctr = 0;
            List<string> csvFiles = Directory.GetFiles(FilePath).ToList();

            foreach (var csvFile in csvFiles)
            {
                var lines = File.ReadAllLines(csvFile).ToList();

                if (lines.Count == 0)
                    continue;

                var headerCols = lines[0].Split(',');
                foreach (var mat in maturityCodes)
                    if (!headerCols.Contains(mat))
                        lines[0] += $",{mat}";

                File.WriteAllLines(csvFile, lines);
                Console.WriteLine($"Updated: {Path.GetFileName(csvFile)}... {++ctr}/{csvFiles.Count}");
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

        public CSVTicker Clone()
        {
            return new CSVTicker
            {
                Ticker = this.Ticker,
                Date = this.Date,
                Maturities = new Dictionary<string, double>(this.Maturities)
            };
        }
    }

}
