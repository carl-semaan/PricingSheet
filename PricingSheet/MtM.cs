using BBGWrapper.Responses.MessagePieces;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Windows.Forms;
using static PricingSheet.Flux;
using static PricingSheet.MtM;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using PricingSheet.Models;
using PricingSheet.Readers;
using PricingSheet.Bloomberg;

namespace PricingSheet
{
    public partial class MtM
    {
        public static MtM MtMInstance { get; private set; }
        public Task FilesLoaded => _filesLoadedTcs.Task;
        public BlockData InstrumentDisplayBlock;
        public SheetUniverse MtMSheetUniverse = new SheetUniverse();
        public List<CSVTicker> CSVdata;

        private SheetDisplay SheetDisplay;
        private TaskCompletionSource<bool> _filesLoadedTcs = new TaskCompletionSource<bool>();

        private readonly object _matrixLock = new object();

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
        public async Task RunInitialization()
        {
            var interopSheet = Globals.ThisWorkbook.Worksheets["Sheet2"];
            var vstoSheet = Globals.Factory.GetVstoObject(interopSheet);

            // Initializing Sheet
            SheetInitialization sheetInitialization = new SheetInitialization(
                vstoSheet,
                "MtM",
                true,
                FreezeRow: 1,
                FreezeColumn: 5
               );

            sheetInitialization.Run();

            // Initializing Data 
            List<RowData> rowData = new List<RowData>();
            List<ColumnData> columnData = new List<ColumnData>();

            JSONReader jsonReader = new JSONReader(Constants.PricingSheetFolderPath, Constants.JSONFileName);

            MtMSheetUniverse.Instruments = jsonReader.LoadClass<Instruments>(nameof(Instruments));
            MtMSheetUniverse.Maturities = jsonReader.LoadClass<Maturities>(nameof(Maturities));

            // Setting the Headers
            List<DataCell> headers = GetHeaders();
            rowData.Add(new RowData(1, 1, headers));

            // Running the headers by cell to apply the formatting
            SheetDisplay = new SheetDisplay(vstoSheet, Rows: rowData);
            SheetDisplay.RunCell();
            rowData.Clear();

            // Minimizing the Older Maturities
            for (int i = 0; i < MtMSheetUniverse.Maturities.Count; i++)
            {
                bool isOutdated = DateTime.ParseExact(MtMSheetUniverse.Maturities[i].Maturity.ToString(), "yyyyMM", CultureInfo.InvariantCulture).Year < DateTime.Now.Year;
                if (isOutdated)
                {
                    Excel.Range colRange = interopSheet.Columns[6 + i];
                    colRange.ColumnWidth = 0;
                }
            }

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

            // Fetch Spot values for underlying and calculate yield
            Task.Run(() => LoadAndDisplay(jsonReader));

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
                bool isOutdated = DateTime.ParseExact(mat.Maturity.ToString(), "yyyyMM", CultureInfo.InvariantCulture).Year <= DateTime.Now.Year;
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
            Ribbons.Ribbon.RibbonInstance?.SetStatus(dbStatus: "Loading...");
            Stopwatch sw = Stopwatch.StartNew();
            CSVdata = await reader.LoadAllTickersAsync(MtMSheetUniverse.Instruments.Select(x => x.Ticker));
            sw.Stop();

            if (CSVdata.Any())
            {
                lock (_matrixLock)
                {
                    foreach (var tickerData in CSVdata)
                    {
                        InstrumentDisplayBlock.UpdateMatrix(tickerData.Ticker, "Last Update", tickerData.Date);
                        foreach (var mat in tickerData.Maturities)
                        {
                            InstrumentDisplayBlock.UpdateMatrix(tickerData.Ticker, string.Concat(mat.Key[0], mat.Key[2]), mat.Value);
                        }
                    }
                    SheetDisplay.RunBlock();
                }
                Ribbons.Ribbon.RibbonInstance?.SetStatus(dbStatus: "Loaded");
            }
            else
                Ribbons.Ribbon.RibbonInstance?.SetStatus(dbStatus: "Failed");

            SignalFilesLoaded();
        }

        public void SignalFilesLoaded()
        {
            _filesLoadedTcs.TrySetResult(true);
        }

        private async Task LoadAndDisplay(JSONReader reader)
        {
            Ribbons.Ribbon.RibbonInstance?.SetStatus(spotStatus: "Loading...");
            List<LastPriceLoad> lastLoad = reader.LoadClass<LastPriceLoad>(nameof(LastPriceLoad));

            List<UnderlyingSpot> rawResponse;
            if (lastLoad.Select(x => x.LastLoad).FirstOrDefault() < DateTime.Today)
                rawResponse = await LoadBloombergPrices(reader);
            else
                rawResponse = LoadSavedPrices(reader);

            var response = rawResponse.ToDictionary(x => x.Underlying, x => x);

            string yieldMaturity = $"Z{DateTime.Now.AddYears(2).Year % 10}";
            double pivotMaturity;

            await FilesLoaded;

            lock (_matrixLock)
            {
                foreach (Instruments instr in MtMSheetUniverse.Instruments)
                {
                    try
                    {
                        if (response.TryGetValue(instr.Underlying, out var res))
                        {
                            if (res.Value == null)
                                continue;

                            InstrumentDisplayBlock.UpdateMatrix(instr.Ticker, "Spot", res.Value);
                            pivotMaturity = Convert.ToDouble(InstrumentDisplayBlock.GetValue(instr.Ticker, yieldMaturity));
                            InstrumentDisplayBlock.UpdateMatrix(instr.Ticker, "Yield", $"{(pivotMaturity / res.Value) * 100}%");
                        }
                        else
                        {
                            Debug.WriteLine($"No value for {instr.Underlying}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error processing {instr.Underlying}: {ex.Message}");
                    }
                }
                SheetDisplay.RunBlock();
            }

            if (response.Any())
                Ribbons.Ribbon.RibbonInstance?.SetStatus(spotStatus: "Loaded");
            else
                Ribbons.Ribbon.RibbonInstance?.SetStatus(spotStatus: "Failed");
        }

        private async Task<List<UnderlyingSpot>> LoadBloombergPrices(JSONReader reader)
        {
            BloombergDataRequest dataRequest = new BloombergDataRequest(MtMInstance, MtMSheetUniverse.Instruments.Select(x => x.Underlying).Distinct().ToList(), new List<string>() { "PX_CLOSE_1D" });

            Stopwatch sw = Stopwatch.StartNew();
            var rawResponse = await dataRequest.FetchUlSpot();
            sw.Stop();

            List<UnderlyingSpot> response = rawResponse.Select(x => new UnderlyingSpot(x.Underlying, x.Value)).ToList();

            if (response.Count == 0)
                return response;

            JSONContent content = new JSONContent();
            content.Instruments = reader.LoadClass<Instruments>(nameof(Instruments));
            content.Maturities = reader.LoadClass<Maturities>(nameof(Maturities));
            content.Fields = reader.LoadClass<Fields>(nameof(Fields));
            content.LastPriceLoad = new List<LastPriceLoad> { new LastPriceLoad(DateTime.Today) };
            content.UnderlyingSpot = response;

            reader.SaveJSON<JSONContent>(content);

            return response;
        }

        private List<UnderlyingSpot> LoadSavedPrices(JSONReader reader) => reader.LoadClass<UnderlyingSpot>(nameof(UnderlyingSpot));
        #endregion

        #region Sheet Update
        public async void RefreshSheet()
        {
            // Clear Sheet
            List<int> ExceptionCol = new List<int>() { 0 };
            InstrumentDisplayBlock.ClearMatrix(new List<int>(), ExceptionCol);

            // Display Empty Block
            SheetDisplay.RunBlock();

            // Fetch Tickers Data
            CSVReader csvReader = new CSVReader(Constants.TickersDBFolderPath);
            JSONReader jsonReader = new JSONReader(Constants.PricingSheetFolderPath, Constants.JSONFileName);
            Task.Run(async () =>
            {
                await LoadAndDisplay(csvReader);
                LoadAndDisplay(jsonReader);
            });
        }
        #endregion
    }
}
