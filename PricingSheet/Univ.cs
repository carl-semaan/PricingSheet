using DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using PricingSheet.Alerts;
using PricingSheet.Models;
using PricingSheet.Readers;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using MelanionMailer;

namespace PricingSheet
{
    public partial class Univ
    {
        public static Univ UnivInstance { get; private set; }
        public BlockGrid Grid { get; set; }

        private SheetUniverse UnivSheetUniverse = new SheetUniverse();
        private SheetDisplay SheetDisplay;
        private ConcurrentQueue<Alert> Alerts = new ConcurrentQueue<Alert>();
        private readonly HashSet<Alert> _activeAlerts = new HashSet<Alert>();
        private readonly object _matrixLock = new object();
        private readonly object _alertsLock = new object();

        private void Sheet1_Startup(object sender, System.EventArgs e)
        {
            UnivInstance = this;
            RunInitialization();
        }

        private void Sheet1_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Sheet1_Startup);
            this.Shutdown += new System.EventHandler(Sheet1_Shutdown);
        }

        #endregion

        #region Sheet Initialization
        public void RunInitialization()
        {
            var interopSheet = Globals.ThisWorkbook.Worksheets["Sheet1"];
            var vstoSheet = Globals.Factory.GetVstoObject(interopSheet);

            // Initializing Sheet
            SheetInitialization sheetInitialization = new SheetInitialization(
                vstoSheet,
                "Univ",
                true,
                FreezeRow: 1,
                FreezeColumn: 1
               );

            sheetInitialization.Run();

            // Initializing Data
            List<RowData> rowData = new List<RowData>();
            List<ColumnData> columnData = new List<ColumnData>();

            JSONReader jsonReader = new JSONReader(Constants.PricingSheetFolderPath, Constants.JSONFileName);

            UnivSheetUniverse.Instruments = jsonReader.LoadClass<Instruments>(nameof(Instruments));
            UnivSheetUniverse.Maturities = jsonReader.LoadClass<Maturities>(nameof(Maturities)).Where(x => x.Flux).ToList();
            UnivSheetUniverse.Fields = jsonReader.LoadClass<Fields>(nameof(Fields));

            // Setting up Matrix Dimensions
            (int width, int height) = GetMatrixDimensions();

            // Setting up the columns and rows headers
            rowData.Add(new RowData(1, 3, GetColHeaders(width)));
            columnData.Add(new ColumnData(2, 1, GetRowHeaders(height)));

            // Building the Display Grid
            Grid = BuildDisplayMatrix(width, height);

            // Display Sheet Values
            SheetDisplay = new SheetDisplay(vstoSheet, columnData, rowData, Grid: Grid);
            SheetDisplay.RunDisplay();

            // Launch Alerts
            LaunchAlerts();
        }

        private (int width, int height) GetMatrixDimensions()
        {
            if (UnivSheetUniverse.Instruments.Count == 0)
                return (0, 0);

            int height = (int)Math.Sqrt(UnivSheetUniverse.Instruments.Count);
            int width = (int)Math.Ceiling((double)UnivSheetUniverse.Instruments.Count / height);

            return (width, height);
        }

        private List<DataCell> GetColHeaders(int width)
        {
            List<DataCell> colHeaders = new List<DataCell>();
            for (int i = 0; i < width; i++)
            {
                foreach (var field in UnivSheetUniverse.Fields)
                {
                    colHeaders.Add(new DataCell(field.Field, IsBold: true, IsCentered: true));
                }
                colHeaders.Add(new DataCell("MtM", IsBold: true, IsCentered: true));
                colHeaders.Add(new DataCell("", IsBold: true, IsCentered: true));
            }
            return colHeaders;
        }

        private List<DataCell> GetRowHeaders(int height)
        {
            List<DataCell> rowHeaders = new List<DataCell>();
            for (int i = 0; i < height; i++)
            {
                foreach (var mat in UnivSheetUniverse.Maturities)
                {
                    rowHeaders.Add(new DataCell(mat.MaturityCode, IsBold: true, IsCentered: true));
                }
            }
            return rowHeaders;
        }

        private BlockGrid BuildDisplayMatrix(int width, int height)
        {
            List<string> cols =
                new[] { "Ticker" }
                    .Concat(UnivSheetUniverse.Fields.Select(x => x.Field))
                    .Concat(new[] { "MtM" })
                    .ToList();

            List<string> rows = UnivSheetUniverse.Maturities.Select(x => x.MaturityCode).ToList();

            BlockGrid grid = new BlockGrid()
            {
                Blocks = new List<BlockData>(),
                StartRow = 2,
                StartColumn = 2,
                Width = width,
                Height = height,
                GridMap = new Dictionary<string, BlockData>()
            };

            int offsetCol = 0; int offsetRow = 0; int instrumentCtr = 0;
            for (int i = 0; i < height && instrumentCtr < UnivSheetUniverse.Instruments.Count; i++)
            {
                for (int j = 0; j < width && instrumentCtr < UnivSheetUniverse.Instruments.Count; j++)
                {
                    string ticker = UnivSheetUniverse.Instruments[instrumentCtr].Ticker;
                    BlockData newBlock = new BlockData(grid.StartRow + offsetRow, grid.StartColumn + offsetCol, rows, cols, HasBorders: true);

                    foreach (var maturity in rows)
                        newBlock.UpdateMatrix(maturity, cols[0], ticker);

                    grid.Blocks.Add(newBlock);
                    if (!grid.GridMap.ContainsKey(ticker))
                        grid.GridMap.Add(ticker, newBlock);

                    offsetCol += cols.Count;
                    instrumentCtr++;
                }
                offsetCol = 0;
                offsetRow += rows.Count;
            }

            return grid;
        }
        #endregion

        #region Sheet Auto Display Update
        private System.Windows.Forms.Timer uiTimer = new System.Windows.Forms.Timer();

        // Auto update method will be called in the Flux sheet to trigger the updates simultaneously to not cause excessive flickering or freeze of the Excel UI
        public void AutoUpdate()
        {
            try
            {
                List<BlockData> dirtyBlocks;

                lock (_matrixLock)
                {
                    dirtyBlocks = Grid.Blocks
                                      .Where(b => b.DirtyFlag)
                                      .ToList();

                    foreach (var block in dirtyBlocks)
                        block.DirtyFlag = false;
                }

                if (dirtyBlocks.Count > 0)
                {
                    Stopwatch sw = Stopwatch.StartNew();
                    SheetDisplay.RunDirtyBlocks(dirtyBlocks);
                    sw.Stop();
                }

                FindAlerts();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error: {ex}");
            }
        }

        public void UpdateMatrixSafe(string instrument, string field, object value)
        {
            string[] parts = instrument.Split('=');

            string maturity = parts[1].Split(' ')[0];
            string ticker = parts[0];

            BlockData target = Grid.GridMap[ticker];
            target.DirtyFlag = true;

            object MtMvalue = MtM.MtMInstance.InstrumentDisplayBlock.GetValue(ticker, maturity);

            lock (_matrixLock)
            {
                target.UpdateMatrix(maturity, field, value);
                target.UpdateMatrix(maturity, "MtM", MtMvalue);
            }
        }

        public void UpdateSubscriptions(List<Maturities> newMaturities)
        {
            lock (_matrixLock)
            {
                UnivSheetUniverse.Maturities = newMaturities.Where(x => x.Flux).ToList();
            }

            // Rebuild the display grid
            Grid.ClearGrid();
            Grid.Blocks.ForEach(x => x.DirtyFlag = true);

            // Clear the alerts
            Alerts = new ConcurrentQueue<Alert>();
            _activeAlerts.Clear();
        }
        #endregion

        #region Sheet Alert System
        public void FindAlerts()
        {
            Stopwatch sw = Stopwatch.StartNew();

            if (Ribbons.Ribbon.RibbonInstance == null || !Ribbons.Ribbon.RibbonInstance.Alerts.Checked)
                return;

            lock (_matrixLock)
            {
                foreach (var block in Grid.Blocks)
                {
                    string instrument = block.GetValue(block.Rows[0], "Ticker") as string;
                    string underlying = UnivSheetUniverse.Instruments.FirstOrDefault(x => x.Ticker == instrument)?.Underlying;

                    foreach (var row in block.Rows)
                    {
                        double askPrice = block.GetValue(row, "ASK") as double? ?? double.NaN;
                        double bidPrice = block.GetValue(row, "BID") as double? ?? double.NaN;
                        double MtMPrice = block.GetValue(row, "MtM") as double? ?? double.NaN;

                        lock (_alertsLock)
                        {
                            if (!double.IsNaN(askPrice) && !double.IsNaN(MtMPrice) && askPrice < MtMPrice)
                            {
                                Alert alert = new Alert(instrument, underlying, row, "ASK", Alert.AlertCondition.LessThan);
                                if (_activeAlerts.Add(alert))
                                    Alerts.Enqueue(alert);
                            }
                            else
                            {
                                _activeAlerts.Remove(new Alert(instrument, underlying, row, "ASK", Alert.AlertCondition.LessThan));
                            }

                            if (!double.IsNaN(bidPrice) && !double.IsNaN(MtMPrice) && bidPrice > MtMPrice)
                            {
                                Alert alert = new Alert(instrument, underlying, row, "BID", Alert.AlertCondition.GreaterThan);
                                if (_activeAlerts.Add(alert))
                                    Alerts.Enqueue(alert);
                            }
                            else
                            {
                                _activeAlerts.Remove(new Alert(instrument, underlying, row, "BID", Alert.AlertCondition.GreaterThan));
                            }
                        }
                    }
                }
            }
            sw.Stop();
        }

        public void LaunchAlerts()
        {
            Task.Run(() =>
            {
                SpeechAlerts speechAlerts = new SpeechAlerts();

                while (true)
                {
                    if (Ribbons.Ribbon.RibbonInstance == null || !Ribbons.Ribbon.RibbonInstance.Alerts.Checked || !Alerts.TryDequeue(out Alert alert))
                    {
                        Thread.Sleep(100);
                        continue;
                    }

                    speechAlerts.Speak(alert.ToString());
                    Task.Run(() => Mailer.SendMailHtml(alert.ToString(), "", alert.ToString(), Constants.Emails));

                    lock (_alertsLock)
                    {
                        if (_activeAlerts.Contains(alert))
                            Alerts.Enqueue(alert);
                    }

                    Thread.Sleep(3500);
                }
            });
        }

        public void ClearAlerts()
        {
            lock (_alertsLock)
            {
                Alerts = new ConcurrentQueue<Alert>();
                _activeAlerts.Clear();
            }
        }
        #endregion
    }
}
