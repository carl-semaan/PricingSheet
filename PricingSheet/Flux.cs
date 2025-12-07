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
        private void Sheet3_Startup(object sender, System.EventArgs e)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
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

            JSONReader reader = new JSONReader(@"G:\Shared drives\Arbitrage\Tools\9.Pricing Sheets", "PricingSheetData.json");
            List<Instruments> instruments = reader.LoadClass<Instruments>(nameof(Instruments));
            List<Maturities> maturities = reader.LoadClass<Maturities>(nameof(Maturities));
            List<Fields> fields = reader.LoadClass<Fields>(nameof(Fields));

            // Merging Cells
            for (int i = 0; i < maturities.Count * 2; i += 2)
            {
                new CellMerge(2, 2, i + 4, i + 5).Run(vstoSheet);
                new CellMerge(1, 1, i + 4, i + 5).Run(vstoSheet);
            }

            // Setting the Upper Headers
            (List<DataCell> maturityCodes, List<DataCell> maturitiesString) = UpperHeaders(maturities, 2, 4);
            rowData.Add(new RowData(1, 4, maturityCodes));
            rowData.Add(new RowData(2, 4, maturitiesString));

            new SheetDisplay(vstoSheet, columnData, rowData).Run();
            rowData.Clear();

            // Setting the Headers
            List<DataCell> Headers = GetHeaders(maturities, fields);
            List<RowData> Rows = new List<RowData>();
            rowData.Add(new RowData(3, 1, Headers));

            // Setting the Data
            columnData.Add(new ColumnData(4, 1, instruments.Select(x => new DataCell(x.Ticker, IsBold: true, IsCentered: true)).ToList()));
            columnData.Add(new ColumnData(4, 2, instruments.Select(x => new DataCell(x.Underlying, IsCentered: true)).ToList()));
            columnData.Add(new ColumnData(4, 3, instruments.Select(x => new DataCell(x.ShortName, IsCentered: true)).ToList()));
            columnData.Add(new ColumnData(4, 4 + maturities.Count * 2, instruments.Select(x => new DataCell(x.ExchangeCode, IsCentered: true)).ToList()));
            columnData.Add(new ColumnData(4, 5 + maturities.Count * 2, instruments.Select(x => new DataCell(x.Currency, IsCentered: true)).ToList()));

            // Setting the Display Block
            InstrumentDisplayBlock = new BlockData(4, 4, instruments.Select(x => x.Ticker).ToList(), maturities.SelectMany(m => fields.Select(f => $"{m.MaturityCode}_{f.Field}")).ToList());

            // Display Sheet Values
            SheetDisplay = new SheetDisplay(vstoSheet, columnData, rowData, InstrumentDisplayBlock);
            SheetDisplay.RunDisplay();

            // Initialize Column and Row Maps
            InitializeDictionaries(interopSheet, maturities.Select(x => x.MaturityCode).ToList(), fields.Select(x => x.Field).ToList(), instruments.Select(x => x.Ticker).ToList());

            // Launch Bloomberg Pipeline
            BloombergPipeline pipeline = new BloombergPipeline(
                this,
                vstoSheet,
                instruments,
                maturities.Select(x => x.MaturityCode).ToList(),
                fields.Select(x => x.Field).ToList()
            );
            Task.Run(() => pipeline.Launch(BloombegCts.Token));

            // Launch Auto Display Update
            StartAutoUpdate();

            //SheetButton sheetButton = new SheetButton(
            //    "Say Hello",
            //    1,
            //    1,
            //    "Blue",
            //    () => System.Windows.Forms.MessageBox.Show("Welcome to the new Pricing Sheet!!!"));

            //List<SheetButton> ButtonsList = new List<SheetButton>();
            //ButtonsList.Add(sheetButton);
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

        #region SheetData
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

        internal class Maturities
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

        internal class Fields
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

        public void StartAutoUpdate()
        {
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
        #endregion
    }
}
