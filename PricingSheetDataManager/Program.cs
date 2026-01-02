using PricingSheet.Bloomberg;
using PricingSheetCore.Models;
using PricingSheetDataManager.Eurex;
using PricingSheetDataManager.Euronext;
using System.Diagnostics;
using System.Threading.Tasks;
using PricingSheetCore;
using DocumentFormat.OpenXml.Wordprocessing;
using PricingSheetCore.Readers;

namespace PricingSheetDataManager
{
    public class Program
    {
        public static async Task Main(string[] args)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();

            // Fetching the live instruments from the exchanges
            Stopwatch sw = Stopwatch.StartNew();

            List<EurexInstruments> eurexInstruments = new List<EurexInstruments>();
            List<EuronextInstruments> euronextInstruments = new List<EuronextInstruments>();
            for (int i = 0; i < Constants.MaxAttempts; i++)
            {
                eurexInstruments = await EurexData.FetchEurexInstruments();
                euronextInstruments = await EuronextData.FetchEuronextInstruments();

                if (eurexInstruments.Count > 0 && euronextInstruments.Count > 0)
                    break;
                else if (i == Constants.MaxAttempts - 1)
                {
                    Console.WriteLine("Maximum attempts reached. Exiting application.");
                    Environment.Exit(1);
                }
                else
                    Console.WriteLine($"Attempt {i + 1} failed. Retrying...");
            }

            sw.Stop();
            Console.WriteLine($"Data fetching and parsing completed in {sw.ElapsedMilliseconds / 1000} seconds.");

            // Unifying the instruments into a single list
            List<Instruments> ListedInstruments = GetAllInstruments(eurexInstruments, euronextInstruments);

            // Fetch the ticker data from Bloomberg
            ListedInstruments = await FetchBloombergData(ListedInstruments);

            // Load previous data from JSON file
            JSONReader reader = new JSONReader(Constants.PricingSheetFolderPath, Constants.JSONFileName);

            List<Maturities> maturities = reader.LoadClass<Maturities>(nameof(Maturities));
            List<Fields> fields = reader.LoadClass<Fields>(nameof(Fields));
            List<LastPriceLoad> lastPriceLoad = reader.LoadClass<LastPriceLoad>(nameof(LastPriceLoad));
            List<UnderlyingSpot> underlyingSpot = reader.LoadClass<UnderlyingSpot>(nameof(UnderlyingSpot));

            // Update maturities
            maturities = UpdateMaturities(maturities);

            // Add missing csv files to database

            // Save changes to JSON file
            JSONContent JsonContent = new JSONContent(ListedInstruments, maturities, fields, lastPriceLoad, underlyingSpot);
            reader.SaveJSON(JsonContent);

            stopwatch.Stop();
            Console.WriteLine("Pricing Sheet was updated successfully");
            Console.WriteLine($"Update completed in: {stopwatch.ElapsedMilliseconds / 1000} seconds.");
        }

        private static List<Instruments> GetAllInstruments(List<EurexInstruments> eurexInstruments, List<EuronextInstruments> euronextInstruments)
        {
            List<Instruments> ListedInstruments = new List<Instruments>();

            foreach (var instrument in eurexInstruments)
                ListedInstruments.Add(
                    new Instruments(
                        instrument.ProductID,
                        string.Empty,
                        string.Empty,
                        exchangeRegionMapping[Exchanges.Eurex][Region.NoRegion],
                        "Equity",
                        string.Empty,
                        string.Empty,
                        "Eurex"
                        )
                    );

            foreach (var instrument in euronextInstruments)
                ListedInstruments.Add(
                    new Instruments(
                        instrument.Code,
                        string.Empty,
                        string.Empty,
                        exchangeRegionMapping[Exchanges.Euronext][(Region)Enum.Parse(typeof(Region), instrument.Location)],
                        "Equity",
                        string.Empty,
                        string.Empty,
                        "Euronext"
                        )
                    );

            return ListedInstruments;
        }

        private static async Task<List<Instruments>> FetchBloombergData(List<Instruments> ListedInstruments)
        {
            Console.WriteLine("Fetching Bloomberg data...");
            List<string> DataFieldsTicker = new List<string>() { "OPT_UNDL_TICKER", "CRNCY", "ICB_SUPERSECTOR_NAME" };
            BloombergDataRequest tickerRequest = new BloombergDataRequest(ListedInstruments.Select(x => x.GetGenericRtCode()).ToList(), DataFieldsTicker);

            var response = await tickerRequest.FetchInstrument();

            foreach (var r in response.ToList())
            {
                var target = ListedInstruments.Where(x => x.GetGenericRtCode() == r.Ticker).FirstOrDefault();
                target.Underlying = r.Underlying;
                target.Currency = r.Currency;
                target.ICBSuperSectorName = r.ICBSuperSectorName;
            }

            List<string> DataFieldsUnderlying = new List<string>() { "SHORT_NAME" };
            BloombergDataRequest underlyingRequest = new BloombergDataRequest(ListedInstruments.Where(x => x.Underlying != null).Select(x => x.GetUlRtCode()).Distinct().ToList(), DataFieldsUnderlying);

            var response2 = await underlyingRequest.FetchInstrument();

            foreach (var i in ListedInstruments)
                i.ShortName = response2.Where(r => r.Ticker == i.GetUlRtCode()).Select(r => r.ShortName).FirstOrDefault() ?? string.Empty;

            return ListedInstruments;
        }

        private static List<Maturities> UpdateMaturities(List<Maturities> maturities)
        {
            Console.WriteLine("Updating maturities...");
            int currentYear = DateTime.Now.Year;

            // Deactivating past maturities
            foreach (var mat in maturities)
            {
                int year = int.Parse(mat.MaturityCode.Substring(1, 2));
                if (year < currentYear % 100)
                {
                    mat.Flux = false;
                    mat.Active = false;
                }
            }

            // Adding missing maturities up to the defined number of years ahead
            List<Maturities> missingMaturities = new List<Maturities>();
            for (int i = 0; i < Constants.MaturitiesAhead; i++)
            {
                if (!maturities.Any(m => m.MaturityCode == $"Z{(i + currentYear) % 100}"))
                {
                    Maturities newMat = new Maturities((currentYear + i) * 100 + 12, $"Z{(currentYear + i) % 100}", true, false);
                    maturities.Add(newMat);
                    missingMaturities.Add(newMat);
                }
            }

            // Add missing maturities to the CSV files
            if(missingMaturities.Count > 0)
            {
                Console.WriteLine("Adding missing maturities to CSV files...");
                CSVReader csvReader = new CSVReader(Constants.TickersDBFolderPath);
                csvReader.AddMaturities(missingMaturities);
            }

            return maturities;
        }

        public static Dictionary<Exchanges, Dictionary<Region, string>> exchangeRegionMapping = new Dictionary<Exchanges, Dictionary<Region, string>>()
            {
                {Exchanges.Eurex, new Dictionary<Region, string>()
                    {
                        { Region.NoRegion, "GR" }
                    }
                },
                { Exchanges.Euronext, new Dictionary<Region, string>()
                    {
                        { Region.Paris, "FP" },
                        { Region.Amsterdam, "NA" },
                        { Region.Milan, "IM" },
                        { Region.Brussels, "BB" },
                        { Region.Oslo, "NO" },
                        { Region.Lisbon, "BD" }
                    }
                },
                { Exchanges.ICE, new Dictionary<Region, string>()
                    {
                        { Region.NoRegion, "LN" }
                    }
                }
            };

        public enum Exchanges
        {
            Eurex,
            Euronext,
            ICE
        }

        public enum Region
        {
            Paris,
            Amsterdam,
            Milan,
            Brussels,
            Oslo,
            Lisbon,
            NoRegion
        }
    }
}