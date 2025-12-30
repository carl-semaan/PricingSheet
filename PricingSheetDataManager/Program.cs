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
            Dictionary<Exchanges, Dictionary<Region, string>> exchangeRegionMapping = new Dictionary<Exchanges, Dictionary<Region, string>>()
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

            Stopwatch sw = Stopwatch.StartNew();

            List<EurexInstruments> eurexInstruments = await EurexData.FetchEurexInstruments();
            List<EuronextInstruments> euronextInstruments = await EuronextData.FetchEuronextInstruments();

            sw.Stop();
            Console.WriteLine($"Data fetching and parsing completed in {sw.ElapsedMilliseconds / 1000} seconds.");
        }
    }

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