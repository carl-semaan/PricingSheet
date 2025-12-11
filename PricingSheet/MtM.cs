using Microsoft.VisualStudio.Tools.Applications.Runtime;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static PricingSheet.Flux;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace PricingSheet
{
    public partial class MtM
    {
        public static MtM MtMInstance { get; private set; }

        private SheetUniverse MtMSheetUniverse = new SheetUniverse();
        private SheetDisplay SheetDisplay;
        private BlockData InstrumentDisplayBlock;

        private void Sheet2_Startup(object sender, System.EventArgs e)
        {
            MtMInstance = this;
            RunInitialization();
        }

        private void Sheet2_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Sheet2_Startup);
            this.Shutdown += new System.EventHandler(Sheet2_Shutdown);
        }

        #endregion


        #region Sheet Initialization
        public void RunInitialization()
        {
            var interopSheet = Globals.ThisWorkbook.Worksheets["Sheet2"];
            var vstoSheet = Globals.Factory.GetVstoObject(interopSheet);

            // Initializing Sheet
            SheetInitialization sheetInitialization = new SheetInitialization(
                vstoSheet,
                "MtM",
                true,
                FreezeRow: 1,
                FreezeColumn: 4
               );

            sheetInitialization.Run();

            // Initializing Data 
            List<RowData> rowData = new List<RowData>();
            List<ColumnData> columnData = new List<ColumnData>();

            JSONReader reader = new JSONReader(Constants.PricingSheetFolderPath, Constants.JSONFileName);

            MtMSheetUniverse.Instruments = reader.LoadClass<Instruments>(nameof(Instruments));
            MtMSheetUniverse.Maturities = reader.LoadClass<Maturities>(nameof(Maturities));

            // Setting the Headers
            List<DataCell> headers = GetHeaders();
            rowData.Add(new RowData(1, 1, headers));

            // Running the headers by cell to apply the formatting
            SheetDisplay = new SheetDisplay(vstoSheet, Rows: rowData);
            SheetDisplay.RunCell();
            rowData.Clear();

            // Setting the Data 
            columnData.Add(new ColumnData(2, 1, MtMSheetUniverse.Instruments.Select(x => new DataCell(x.Ticker, IsBold: true, IsCentered: true)).ToList()));
            columnData.Add(new ColumnData(2, 2, MtMSheetUniverse.Instruments.Select(x => new DataCell(x.Underlying, IsBold: true, IsCentered: true)).ToList()));
            columnData.Add(new ColumnData(2, 3, MtMSheetUniverse.Instruments.Select(x => new DataCell(x.Currency, IsBold: true, IsCentered: true)).ToList()));
            columnData.Add(new ColumnData(2, 7 + MtMSheetUniverse.Maturities.Count, MtMSheetUniverse.Instruments.Select(x => new DataCell(x.ICBSuperSectorName, IsBold: true, IsCentered: true)).ToList()));

            // Setting the Display Block
            InstrumentDisplayBlock = new BlockData(StartRow: 2, StartColumn: 4, MtMSheetUniverse.Instruments.Select(i => i.Ticker).ToList(), headers.Skip(3).Take(headers.Count - 4).Select(x => x.Value).ToList());

            // Fetch Tickers Data
            CSVReader csvReader = new CSVReader(Constants.TickersDBFolderPath);
            Task.Run(() => LoadAndDisplay(csvReader));

            // Display Sheet Values
            SheetDisplay = new SheetDisplay(vstoSheet, Columns: columnData, Block: InstrumentDisplayBlock);
            SheetDisplay.RunDisplay();
        }

        private List<DataCell> GetHeaders()
        {
            List<string> headerNames = new List<string>() { "Ticker", "Underlying", "Currency", "Spot", "Last Update" };

            List<DataCell> headers = new List<DataCell>();

            foreach (string header in headerNames)
                headers.Add(new DataCell(header, IsBold: true, IsCentered: true));

            foreach (Maturities mat in MtMSheetUniverse.Maturities)
            {
                bool isOutdated = DateTime.ParseExact(mat.Maturity, "yyyyMM", CultureInfo.InvariantCulture).Year <= DateTime.Now.Year;
                headers.Add(new DataCell(mat.MaturityCode, Color: isOutdated ? "Blue" : "Black", IsBold: true, BgColor: isOutdated ? "" : "LightBlue", IsCentered: true));
            }

            headers.Add(new DataCell("Yield", IsBold: true, IsCentered: true));
            headers.Add(new DataCell("ICB Supersector Name", IsBold: true, IsCentered: true));

            return headers;
        }

        private async Task LoadMaturityValues(CSVReader reader)
        {
            List<string> missingTickers = new List<string>();
            List<CSVTicker> bag = new List<CSVTicker>();

            Stopwatch sw = Stopwatch.StartNew();

            foreach (string ticker in MtMSheetUniverse.Instruments.Select(x => x.Ticker))
            {
                try
                {
                    bag.Add(reader.LoadTickerData(ticker));
                }
                catch
                {
                    missingTickers.Add(ticker);
                }
            }

            sw.Stop();
            Debug.WriteLine($"Loaded {bag.Count} tickers in {sw.ElapsedMilliseconds} ms. Missing tickers: {missingTickers.Count}");
        }

        public async Task LoadAndDisplay(CSVReader reader)
        {
            List<CSVTicker> data = await reader.LoadAllTickersAsync(MtMSheetUniverse.Instruments.Select(x => x.Ticker));

            foreach(var tickerData in data)
            {
                InstrumentDisplayBlock.UpdateMatrix(tickerData.Ticker, "Last Update", tickerData.Date);
                foreach(var mat in tickerData.Maturities)
                {
                    InstrumentDisplayBlock.UpdateMatrix(tickerData.Ticker, string.Concat(mat.Key[0], mat.Key[2]), mat.Value);
                }
            }

            SheetDisplay.RunBlock();
        }
        #endregion


        #region Sheet Data
        public class SheetUniverse
        {
            public List<Instruments> Instruments { get; set; }
            public List<Maturities> Maturities { get; set; }
            public SheetUniverse() { }
            public SheetUniverse(List<Instruments> instruments, List<Maturities> maturities)
            {
                Instruments = instruments;
                Maturities = maturities;
            }
        }

        public class Instruments
        {
            public string Ticker { get; set; }
            public string Underlying { get; set; }
            public string Currency { get; set; }
            public string ICBSuperSectorName { get; set; }


            public Instruments() { }
            public Instruments(string ticker, string underlying, string currency, string ICBSuperSectorName)
            {
                Ticker = ticker;
                Underlying = underlying;
                Currency = currency;
                this.ICBSuperSectorName = ICBSuperSectorName;
            }
        }

        public class Maturities
        {
            public string MaturityCode { get; set; }
            public string Maturity { get; set; }

            public Maturities() { }

            public Maturities(string Maturity, string MaturityCode)
            {
                this.MaturityCode = MaturityCode;
                this.Maturity = Maturity;
            }
        }
        #endregion
    }
}
