using Microsoft.VisualStudio.Tools.Applications.Runtime;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static PricingSheet.Flux;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace PricingSheet
{
    public partial class Flux
    {
        public static Dictionary<(string maturity, string field), int> ColMap = new Dictionary<(string maturity, string field), int>();
        public static Dictionary<string, int> RowMap = new Dictionary<string, int>();
        public CancellationTokenSource BloombegCts = new CancellationTokenSource();
        public BlockData InstrumentDisplayBlock;
        private SheetDisplay SheetDisplay;
        private readonly object _matrixLock = new object();
        private SheetUniverse FluxSheetUniverse = new SheetUniverse();
        public static Flux FluxInstance { get; private set; }
        private void Sheet3_Startup(object sender, System.EventArgs e)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            FluxInstance = this;
            RunInitialization();
            sw.Stop();
        }

        private void Sheet3_Shutdown(object sender, System.EventArgs e)
        {
            BloombegCts.Cancel();
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Sheet3_Startup);
            this.Shutdown += new System.EventHandler(Sheet3_Shutdown);
        }
        #endregion

        #region Sheet Initialization
        private void RunInitialization()
        {
            var interopSheet = Globals.ThisWorkbook.Worksheets["Sheet3"];
            var vstoSheet = Globals.Factory.GetVstoObject(interopSheet);

            // Initializing Sheet with buttons
            SheetInitialization sheetInitialization = new SheetInitialization(
                vstoSheet,
                "Flux",
                true,
                new List<SheetButton>(),
                FreezeRow: 3
            );

            sheetInitialization.Run();

            // Initializing Data
            List<ColumnData> columnData = new List<ColumnData>();
            List<RowData> rowData = new List<RowData>();

            JSONReader reader = new JSONReader(Constants.FolderPath, Constants.JSONFileName);

            FluxSheetUniverse.Instruments = reader.LoadClass<Instruments>(nameof(Instruments));
            FluxSheetUniverse.Maturities = reader.LoadClass<Maturities>(nameof(Maturities));
            FluxSheetUniverse.Fields = reader.LoadClass<Fields>(nameof(Fields));

            // Merging Cells
            for (int i = 0; i < FluxSheetUniverse.Maturities.Count * 2; i += 2)
            {
                new CellMerge(2, 2, i + 4, i + 5).Run(vstoSheet);
                new CellMerge(1, 1, i + 4, i + 5).Run(vstoSheet);
            }

            // Setting the Upper Headers
            (List<DataCell> maturityCodes, List<DataCell> maturitiesString) = UpperHeaders(FluxSheetUniverse.Maturities, 2, 4);
            rowData.Add(new RowData(1, 4, maturityCodes));
            rowData.Add(new RowData(2, 4, maturitiesString));

            new SheetDisplay(vstoSheet, columnData, rowData).Run();
            rowData.Clear();

            // Setting the Headers
            List<DataCell> Headers = GetHeaders(FluxSheetUniverse.Maturities, FluxSheetUniverse.Fields);
            List<RowData> Rows = new List<RowData>();
            rowData.Add(new RowData(3, 1, Headers));

            // Setting the Data
            columnData.Add(new ColumnData(4, 1, FluxSheetUniverse.Instruments.Select(x => new DataCell(x.Ticker, IsBold: true, IsCentered: true)).ToList()));
            columnData.Add(new ColumnData(4, 2, FluxSheetUniverse.Instruments.Select(x => new DataCell(x.Underlying, IsCentered: true)).ToList()));
            columnData.Add(new ColumnData(4, 3, FluxSheetUniverse.Instruments.Select(x => new DataCell(x.ShortName, IsCentered: true)).ToList()));
            columnData.Add(new ColumnData(4, 4 + FluxSheetUniverse.Maturities.Count * 2, FluxSheetUniverse.Instruments.Select(x => new DataCell(x.ExchangeCode, IsCentered: true)).ToList()));
            columnData.Add(new ColumnData(4, 5 + FluxSheetUniverse.Maturities.Count * 2, FluxSheetUniverse.Instruments.Select(x => new DataCell(x.Currency, IsCentered: true)).ToList()));

            // Setting the Display Block
            InstrumentDisplayBlock = new BlockData(4, 4, FluxSheetUniverse.Instruments.Select(x => x.Ticker).ToList(), FluxSheetUniverse.Maturities.SelectMany(m => FluxSheetUniverse.Fields.Select(f => $"{m.MaturityCode}_{f.Field}")).ToList());

            // Display Sheet Values
            SheetDisplay = new SheetDisplay(vstoSheet, columnData, rowData, InstrumentDisplayBlock);
            SheetDisplay.RunDisplay();

            // Initialize Column and Row Maps
            InitializeDictionaries(interopSheet, FluxSheetUniverse.Maturities.Select(x => x.MaturityCode).ToList(), FluxSheetUniverse.Fields.Select(x => x.Field).ToList(), FluxSheetUniverse.Instruments.Select(x => x.Ticker).ToList());

            // Launch Bloomberg Pipeline
            BloombergPipeline pipeline = new BloombergPipeline(
                this,
                vstoSheet,
                FluxSheetUniverse.Instruments,
                FluxSheetUniverse.Maturities.Select(x => x.MaturityCode).ToList(),
                FluxSheetUniverse.Fields.Select(x => x.Field).ToList()
            );
            Task.Run(() => pipeline.Launch(BloombegCts.Token));

            // Launch Auto Display Update
            StartAutoUpdate(BloombegCts.Token);

        }

        private static (List<DataCell>, List<DataCell>) UpperHeaders(List<Maturities> maturities, int row, int column)
        {
            List<DataCell> Codes = new List<DataCell>();

            List<DataCell> Maturities = new List<DataCell>();

            foreach (var mat in maturities)
            {
                Codes.Add(new DataCell(mat.MaturityCode, IsBold: true, IsCentered: true, Offset: 1));
                Maturities.Add(new DataCell(mat.Maturity.ToString(), IsBold: true, IsCentered: true, Offset: 1));
            }

            return (Maturities, Codes);
        }

        private static List<DataCell> GetHeaders(List<Maturities> maturities, List<Fields> fields)
        {
            List<string> headers = new List<string>() { "Ticker", "Underlying", "Short Name" };

            foreach (var mat in maturities)
            {
                foreach (var field in fields)
                {
                    headers.Add($"{field.Field}");
                }
            }

            headers.Add("Exchange Code");
            headers.Add("Currency");

            var newHeaders = headers.Select((h, index) => new DataCell(h, IsBold: true, IsCentered: true)).ToList();

            return newHeaders;
        }

        private static void InitializeDictionaries(
            ExcelInterop.Worksheet sheet,
            List<string> maturityCodes,
            List<string> fields,
            List<string> instruments,
            int maturityRow = 2,
            int fieldRow = 3,
            int instrumentColumn = 1,
            int startingColumn = 4,
            int startingRow = 4
            )
        {
            Dictionary<(string maturity, string field), int> ColDictionary = new Dictionary<(string maturity, string field), int>();

            int lastColumn = startingColumn + maturityCodes.Count * fields.Count - 1;
            for (int col = startingColumn; col <= lastColumn; col += fields.Count)
            {
                string colMaturity = ((sheet.Cells[maturityRow, col] as ExcelInterop.Range)?.Value2 ?? "").ToString().Trim().ToLower();
                for (int j = 0; j < fields.Count; j++)
                {
                    string colField = ((sheet.Cells[fieldRow, col + j] as ExcelInterop.Range)?.Value2 ?? "").ToString().Trim().ToLower();
                    ColDictionary[(colMaturity, colField)] = col + j;
                }
            }
            ColMap = ColDictionary;

            Dictionary<string, int> RowDictionary = new Dictionary<string, int>();
            for (int row = startingRow; row < instruments.Count + startingRow; row++)
            {
                string rowInstrument = ((sheet.Cells[row, instrumentColumn] as ExcelInterop.Range)?.Value2 ?? "").ToString().Trim().ToLower();
                RowDictionary[rowInstrument] = row;
            }
            RowMap = RowDictionary;
        }
        #endregion

        #region Sheet Data
        public class SheetUniverse
        {
            public List<Instruments> Instruments { get; set; }
            public List<Maturities> Maturities { get; set; }
            public List<Fields> Fields { get; set; }
            public SheetUniverse() { }
            public SheetUniverse(List<Instruments> instruments, List<Maturities> maturities, List<Fields> fields)
            {
                Instruments = instruments;
                Maturities = maturities;
                Fields = fields;
            }
        }

        public class Instruments
        {
            public string Ticker { get; set; }
            public string Underlying { get; set; }
            public string ShortName { get; set; }
            public string ExchangeCode { get; set; }
            public string InstrumentType { get; set; }
            public string Currency { get; set; }


            public Instruments() { }
            public Instruments(string ticker, string underlying, string shortName, string exchangeCode, string instrumentType, string currency)
            {
                Ticker = ticker;
                Underlying = underlying;
                ShortName = shortName;
                ExchangeCode = exchangeCode;
                InstrumentType = instrumentType;
                Currency = currency;
            }
        }

        public class Maturities
        {
            public int Maturity { get; set; }
            public string MaturityCode { get; set; }

            public Maturities() { }

            public Maturities(int Maturity, string MaturityCode)
            {
                this.Maturity = Maturity;
                this.MaturityCode = MaturityCode;
            }
        }

        public class Fields
        {
            public string Field { get; set; }

            public Fields() { }

            public Fields(string Field)
            {
                this.Field = Field;
            }
        }
        #endregion

        #region Sheet Auto Display Update
        private System.Windows.Forms.Timer uiTimer = new System.Windows.Forms.Timer();

        public void StartAutoUpdate(CancellationToken token)
        {
            if (token.IsCancellationRequested)
                return;

            uiTimer.Interval = 500;
            uiTimer.Tick += (s, e) =>
            {
                lock (_matrixLock)
                {
                    SheetDisplay.RunBlock();
                }
            };
            uiTimer.Start();
        }

        public void UpdateMatrixSafe(string instrument, string field, object value)
        {
            string[] parts = instrument.Split('=');

            string maturity = parts[1].Split(' ')[0];
            string ticker = parts[0];

            string Maturity_Field = $"{maturity}_{field}";
            lock (_matrixLock)
            {
                InstrumentDisplayBlock.UpdateMatrix(ticker, Maturity_Field, value);
            }
        }

        public void UpdateDisplay(Instruments newInstrument)
        {
            var interopSheet = Globals.ThisWorkbook.Worksheets["Flux"];
            var vstoSheet = Globals.Factory.GetVstoObject(interopSheet);

            FluxSheetUniverse.Instruments.Add(newInstrument);

            // Update Block Data 
            lock (_matrixLock)
            {
                List<string> instruments = InstrumentDisplayBlock.Rows.ToList();
                instruments.Add(newInstrument.Ticker);

                InstrumentDisplayBlock = new BlockData(4, 4, instruments, InstrumentDisplayBlock.Columns);
            }

            // Update Sheet Display
            List<ColumnData> columnData = SheetDisplay.Columns.ToList();
            columnData[0] = new ColumnData(columnData[0].StartRow, columnData[0].Column, columnData[0].Data.Append(new DataCell(newInstrument.Ticker, IsBold: true, IsCentered: true)).ToList());
            columnData[1] = new ColumnData(columnData[1].StartRow, columnData[1].Column, columnData[1].Data.Append(new DataCell(newInstrument.Underlying, IsBold: true, IsCentered: true)).ToList());
            columnData[2] = new ColumnData(columnData[2].StartRow, columnData[2].Column, columnData[2].Data.Append(new DataCell(newInstrument.ShortName, IsBold: true, IsCentered: true)).ToList());
            columnData[3] = new ColumnData(columnData[3].StartRow, columnData[3].Column, columnData[3].Data.Append(new DataCell(newInstrument.ExchangeCode, IsBold: true, IsCentered: true)).ToList());
            columnData[4] = new ColumnData(columnData[4].StartRow, columnData[4].Column, columnData[4].Data.Append(new DataCell(newInstrument.Currency, IsBold: true, IsCentered: true)).ToList());

            SheetDisplay.Columns = columnData;
            SheetDisplay.Block = InstrumentDisplayBlock;
            SheetDisplay.RunDisplay();

            // Cancel Old Bloomberg Pipeline
            BloombegCts.Cancel();

            // Create a new cancelation token
            BloombegCts = new CancellationTokenSource();

            // Update Bloomberg Pipeline
            BloombergPipeline pipeline = new BloombergPipeline(
                this,
                vstoSheet,
                FluxSheetUniverse.Instruments,
                FluxSheetUniverse.Maturities.Select(x => x.MaturityCode).ToList(),
                FluxSheetUniverse.Fields.Select(x => x.Field).ToList()
            );
            Task.Run(() => pipeline.LaunchOfflineTest(BloombegCts.Token));
        }
        #endregion
    }
}
