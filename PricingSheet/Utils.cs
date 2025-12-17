using DocumentFormat.OpenXml.Office2010.PowerPoint;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using ExcelVSTO = Microsoft.Office.Tools.Excel;

namespace PricingSheet
{
    internal class Utils
    {
        public static (int row, int column) FindCellFlux(string maturity, string field, string ticker)
        {
            if (!Flux.ColMap.TryGetValue((maturity.Trim().ToLower(), field.Trim().ToLower()), out int col))
                throw new Exception($"Column not found for: {maturity} - {field}");

            if (!Flux.RowMap.TryGetValue(ticker.Trim().ToLower(), out int row))
                throw new Exception($"Row not found for: {ticker}");

            return (row, col);
        }

    }

    public class Constants
    {
        public const string PricingSheetFolderPath = @"G:\Shared drives\Arbitrage\Tools\9.Pricing Sheets";
        public const string JSONFileName = "PricingSheetData.json";
        public const string TickersDBFolderPath = @"G:\Shared drives\Arbitrage\Tools\9.Pricing Sheets\SSDF Database-Testing";
        public const int TimeoutMS = 5000;
        public const int UiTickInterval = 500;
        public const int ThreadSleep = 1000;
        public const int MaxActiveInstruments = 3500;
    }

}
