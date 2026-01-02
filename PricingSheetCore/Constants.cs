using DocumentFormat.OpenXml.Office2010.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace PricingSheetCore
{
    public class Constants
    {
        public const string PricingSheetFolderPath = @"G:\Shared drives\Arbitrage\Tools\9.Pricing Sheets";
        public const string JSONFileName = "PricingSheetData.json";
        public const string TickersDBFolderPath = @"G:\Shared drives\Arbitrage\Tools\9.Pricing Sheets\SSDF Database-Testing";
        public const int TimeoutMS = 100;
        public const int RequestTimeoutMS = 5000;
        public const int UiTickInterval = 500;
        public const int ThreadSleep = 1000;
        public const int MaxActiveInstruments = 3500;
        public static List<string> Emails = new List<string> { "carl.semaan@melanion.com" };
        // tony.khreich@melanion.com
        // roland.nasr@melanion.com

        public const string EurexURL = @"https://www.eurex.com/ex-en/markets/productSearch";
        public const string EuronextURL = @"https://live.euronext.com/en/products/dividend-stock-futures/list";
        public const int MaxAttempts = 5;
        public const int ScraperTimeoutMS = 120000;
        public const int MaturitiesAhead = 7;
    }
}
