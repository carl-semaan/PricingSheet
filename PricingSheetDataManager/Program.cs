using PricingSheet.Bloomberg;
using PricingSheetCore.Models;
using PricingSheetDataManager.Eurex;
using PricingSheetDataManager.Euronext;
using System.Diagnostics;
using System.Threading.Tasks;

namespace PricingSheetDataManager
{
    public class Program
    {
        public static async Task Main(string[] args)
        {
            // Fetching the live instruments from the exchanges
            Stopwatch sw = Stopwatch.StartNew();

            List<EurexInstruments> eurexInstruments = await EurexData.FetchEurexInstruments();
            List<EuronextInstruments> euronextInstruments = await EuronextData.FetchEuronextInstruments();

            sw.Stop();
            Console.WriteLine($"Data fetching and parsing completed in {sw.ElapsedMilliseconds / 1000} seconds.");

            // Unifying the instruments into a single list
            List<Instruments> ListedInstruments = GetAllInstruments(eurexInstruments, euronextInstruments);

            //Fetch the ticker data from Bloomberg
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
            BloombergDataRequest underlyingRequest = new BloombergDataRequest(ListedInstruments.Select(x => x.Underlying).Distinct().ToList(), DataFieldsUnderlying);

            var response2 = await underlyingRequest.FetchInstrument();

            foreach (var r in response2.ToList())
            {
                var target = ListedInstruments.Where(x => x.Underlying == r.Ticker).FirstOrDefault();
                target.ShortName= r.ShortName;
            }
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