using DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using PricingSheet.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace PricingSheet
{
    public partial class Univ
    {
        public static Univ UnivInstance { get; private set; }

        private SheetUniverse UnivSheetUniverse = new SheetUniverse();

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
            (int height, int width) = GetMatrixDimensions();

            // Setting up the columns and rows headers
            rowData.Add(new RowData(1, 3, GetColHeaders(width)));
            columnData.Add(new ColumnData(2, 1, GetRowHeaders(height)));

            // Building the Display Grid
            BlockGrid Grid = BuildDisplayMatrix(width, height);

            // Display Sheet Values
            SheetDisplay display = new SheetDisplay(vstoSheet, columnData, rowData, Grid: Grid);
            display.RunDisplay();
        }

        private (int height, int width) GetMatrixDimensions()
        {
            if (UnivSheetUniverse.Instruments.Count == 0)
                return (0, 0);

            int height = (int)Math.Sqrt(UnivSheetUniverse.Instruments.Count);
            int width = (int)Math.Ceiling((double)UnivSheetUniverse.Instruments.Count / height);

            return (height, width);
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

                    foreach(var maturity in rows)
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
    }
}
