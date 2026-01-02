using Bloomberglp.Blpapi;
using DocumentFormat.OpenXml.Vml.Office;
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
using ExcelVSTO = Microsoft.Office.Tools.Excel;
using Office = Microsoft.Office.Core;
using PricingSheetCore.Models;
using PricingSheetCore.Readers;
using PricingSheetCore;

namespace PricingSheet
{
    public partial class Flux
    {
        public static Flux FluxInstance { get; private set; }

        public static Dictionary<(string maturity, string field), int> ColMap = new Dictionary<(string maturity, string field), int>();
        public static Dictionary<string, int> RowMap = new Dictionary<string, int>();
        public CancellationTokenSource BloombegCts = new CancellationTokenSource();
        public SheetUniverse FluxSheetUniverse = new SheetUniverse();

        private SheetDisplay SheetDisplay;
        private BlockData InstrumentDisplayBlock;
        private readonly object _matrixLock = new object();
        private ExcelVSTO.Worksheet vstoSheet;

        private void Sheet3_Startup(object sender, System.EventArgs e)
        {
            FluxInstance = this;
            RunInitialization();
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
            vstoSheet = Globals.Factory.GetVstoObject(interopSheet);

            // Initializing Sheet 
            SheetInitialization sheetInitialization = new SheetInitialization(
                vstoSheet,
                "Flux",
                true,
                FreezeRow: 3
            );

            sheetInitialization.Run();

            // Initializing Data
            List<ColumnData> columnData = new List<ColumnData>();
            List<RowData> rowData = new List<RowData>();

            JSONReader reader = new JSONReader(Constants.PricingSheetFolderPath, Constants.JSONFileName);

            FluxSheetUniverse.Instruments = reader.LoadClass<Instruments>(nameof(Instruments));
            FluxSheetUniverse.Maturities = reader.LoadClass<Maturities>(nameof(Maturities)).Where(M => M.Flux).ToList();
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

            // Running the upper headers by cell to apply the formatting
            new SheetDisplay(vstoSheet, Rows: rowData).RunCell();
            rowData.Clear();

            // Setting the Headers
            List<DataCell> Headers = GetHeaders();
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
            LaunchBloomberg();

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

        private List<DataCell> GetHeaders()
        {
            List<string> headers = new List<string>() { "Ticker", "Underlying", "Short Name" };

            foreach (var mat in FluxSheetUniverse.Maturities)
            {
                foreach (var field in FluxSheetUniverse.Fields)
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

        public void LaunchBloomberg()
        {
            BloombergPipeline pipeline = new BloombergPipeline(
                vstoSheet,
                FluxSheetUniverse.Instruments,
                FluxSheetUniverse.Maturities.Where(x => x.Active).Select(x => x.MaturityCode).ToList(),
                FluxSheetUniverse.Fields.Select(x => x.Field).ToList()
            );

            Task.Run(() => pipeline.Launch(BloombegCts.Token));
        }
        #endregion

        #region Sheet Auto Display Update
        private System.Windows.Forms.Timer uiTimer = new System.Windows.Forms.Timer();

        public void StartAutoUpdate(CancellationToken token)
        {
            if (token.IsCancellationRequested)
                return;

            uiTimer.Interval = Constants.UiTickInterval;
            uiTimer.Tick += (s, e) =>
            {
                // Flux Sheet Update
                AutoUpdate();

                // Univ Sheet Update
                Univ.UnivInstance.AutoUpdate();
            };
            uiTimer.Start();
        }

        private void AutoUpdate()
        {
            if (InstrumentDisplayBlock.DirtyFlag)
            {
                InstrumentDisplayBlock.DirtyFlag = false;
                lock (_matrixLock)
                {
                    SheetDisplay.RunBlock();
                }
            }
        }

        public void UpdateMatrixSafe(string instrument, string field, object value)
        {
            InstrumentDisplayBlock.DirtyFlag = true;

            string[] parts = instrument.Split('=');

            string maturity = parts[1].Split(' ')[0];

            int lastDigit = maturity[1] - '0';
            int currentYear = DateTime.Now.Year;
            int currentDigit = currentYear % 10;
            int decade = (currentYear / 10) % 10;

            if (lastDigit < currentDigit)
                decade++;

            maturity = $"{maturity[0]}{decade}{lastDigit}";

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
            InitializeDictionaries(interopSheet, FluxSheetUniverse.Maturities.Select(x => x.MaturityCode).ToList(), FluxSheetUniverse.Fields.Select(x => x.Field).ToList(), FluxSheetUniverse.Instruments.Select(x => x.Ticker).ToList());

            lock (_matrixLock)
            {
                // Update Block Data 
                List<string> instruments = InstrumentDisplayBlock.Rows.ToList();
                instruments.Add(newInstrument.Ticker);

                InstrumentDisplayBlock = new BlockData(4, 4, instruments, InstrumentDisplayBlock.Columns);

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
            }

            // Cancel Old Bloomberg Pipeline
            BloombegCts.Cancel();

            // Create a new cancelation token
            BloombegCts = new CancellationTokenSource();

            // Update Bloomberg Pipeline
            BloombergPipeline pipeline = new BloombergPipeline(
                vstoSheet,
                FluxSheetUniverse.Instruments,
                FluxSheetUniverse.Maturities.Select(x => x.MaturityCode).ToList(),
                FluxSheetUniverse.Fields.Select(x => x.Field).ToList()
            );
            Task.Run(() => pipeline.LaunchOfflineTest(BloombegCts.Token));
        }

        public void UpdateSubscriptions(List<Maturities> newMaturities)
        {
            BloombegCts.Cancel();
            BloombegCts = new CancellationTokenSource();

            Ribbons.Ribbon.RibbonInstance?.SetStatus(bbgStatus: "Pending");

            InstrumentDisplayBlock.ClearMatrix();

            lock (_matrixLock)
            {
                FluxSheetUniverse.Maturities = newMaturities.Where(M => M.Flux).ToList();

                if (!FluxSheetUniverse.Maturities.Where(x => x.Active).Any())
                {
                    // Clear the sheet if no maturities are selected
                    SheetDisplay.RunDisplay();
                    Ribbons.Ribbon.RibbonInstance?.SetActiveSubscription(0);
                    return;
                }
            }

            LaunchBloomberg();
        }
        #endregion
    }
}
