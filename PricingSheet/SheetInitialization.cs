using DocumentFormat.OpenXml.Spreadsheet;
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
    public class SheetInitialization
    {
        public ExcelVSTO.Worksheet Sheet { get; set; }
        public string Name { get; set; }
        public bool ClearOnStartUp { get; set; }
        public List<SheetButton> SheetButtons { get; set; }
        public int FreezeRow { get; set; }
        public int FreezeColumn { get; set; }

        public SheetInitialization() { }

        public SheetInitialization(ExcelVSTO.Worksheet Sheet, string Name, bool ClearOnStartUp, List<SheetButton> SheetButtons, int FreezeRow = 0, int FreezeColumn = 0)
        {
            this.Sheet = Sheet;
            this.Name = Name;
            this.ClearOnStartUp = ClearOnStartUp;
            this.SheetButtons = SheetButtons;
            this.FreezeRow = FreezeRow;
            this.FreezeColumn = FreezeColumn;
        }

        public void Run()
        {
            if (Sheet != null)
            {
                Sheet.Name = Name;

                if (ClearOnStartUp)
                    Sheet.Cells.Clear();

                foreach (var btn in SheetButtons)
                    AddButton(btn);

                FreezePanes();
            }
        }

        private void AddButton(SheetButton btn)
        {
            ExcelInterop.Range cell = Sheet.Cells[btn.Row, btn.Column] as ExcelInterop.Range;

            int left = (int)cell.Left;
            int top = (int)cell.Top;

            var button = Sheet.Controls.AddButton(left, top, btn.Width, btn.Height, btn.Name);
            button.Text = btn.Name;

            button.Click += (s, e) => btn.Action();
        }

        private void FreezePanes()
        {
            if (Sheet == null)
                return;

            if (FreezeRow <= 0 && FreezeColumn <= 0)
                return;

            Sheet.Activate();

            ExcelInterop.Worksheet interopSheet = Sheet.InnerObject;
            var window = interopSheet.Application.ActiveWindow;

            if (FreezeRow > 0)
                window.SplitRow = FreezeRow;

            if (FreezeColumn > 0)
                window.SplitColumn = FreezeColumn;

            window.FreezePanes = true;
        }

    }

    public class SheetButton
    {
        public string Name { get; set; }
        public int Row { get; set; }
        public int Column { get; set; }
        public string Color { get; set; }
        public System.Action Action { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }

        public SheetButton() { }

        public SheetButton(string Name, int Row, int Column, string Color, System.Action Action, int Width = 100, int Height = 30)
        {
            this.Name = Name;
            this.Row = Row;
            this.Column = Column;
            this.Color = Color;
            this.Action = Action;
            this.Width = Width;
            this.Height = Height;
        }
    }

    public class SheetDisplay
    {
        public ExcelVSTO.Worksheet Sheet { get; set; }
        public List<ColumnData> Columns { get; set; }
        public List<RowData> Rows { get; set; }
        public BlockData Block { get; set; }

        public SheetDisplay(ExcelVSTO.Worksheet Sheet, List<ColumnData> Columns, List<RowData> Rows, BlockData Block = null)
        {
            this.Sheet = Sheet;
            this.Columns = Columns;
            this.Rows = Rows;
            this.Block = Block;
        }

        public void RunDisplay(bool batch = true)
        {
            if (batch)
                RunBatch();
            else
                Run();

            if (Block != null)
                RunBlock();
        }

        public void RunBlock()
        {
            if (Sheet == null || Block == null || Block.Data == null)
                return;

            int rows = Block.Data.GetLength(0);
            int cols = Block.Data.GetLength(1);

            if (rows == 0 || cols == 0)
                return;

            try
            {
                ExcelInterop.Application app = Sheet.Application;
                bool prevUpdate = app.ScreenUpdating;
                app.ScreenUpdating = false;

                try
                {
                    ExcelInterop.Range topLeft = Sheet.Cells[Block.StartRow, Block.StartColumn];
                    ExcelInterop.Range bottomRight = Sheet.Cells[Block.StartRow + rows - 1, Block.StartColumn + cols - 1];
                    ExcelInterop.Range writeRange = Sheet.Range[topLeft, bottomRight];

                    writeRange.Value2 = Block.Data;
                }
                finally
                {
                    app.ScreenUpdating = prevUpdate;
                }
            }
            catch (Exception) { }
        }

        public void RunBatch()
        {
            foreach (ColumnData col in Columns)
            {
                if (col.Data.Count == 0)
                    continue;

                DisplayBatchColumn(Sheet, col.Data, col.StartRow, col.Column);
            }

            foreach (RowData row in Rows)
            {
                if (row.Data.Count == 0)
                    continue;

                DisplayBatchRow(Sheet, row.Data, row.Row, row.StartColumn);
            }
        }

        public void Run()
        {
            foreach (ColumnData col in Columns)
            {
                for (int i = 0; i < col.Data.Count; i++)
                {
                    DisplayCell(Sheet, col.Data[i], col.StartRow + i, col.Column);
                }
            }

            foreach (RowData row in Rows)
            {
                int offSet = 0;
                for (int i = 0; i < row.Data.Count; i++)
                {
                    DisplayCell(Sheet, row.Data[i], row.Row, row.StartColumn + i + offSet);
                    if (row.Data[i].Offset != 0)
                        offSet += row.Data[i].Offset;
                }
            }
        }

        public static void DisplayBatchRow(ExcelVSTO.Worksheet sheet, List<DataCell> dataCells, int row, int StartColumn)
        {
            int n = dataCells.Count;
            object[,] values = new object[1, n];

            for (int i = 0; i < n; i++)
                values[0, i] = dataCells[i].Value;

            var range = sheet.Range[
                sheet.Cells[row, StartColumn],
                sheet.Cells[row, StartColumn + n - 1]
            ];


            range.Value2 = values;
            if (!string.IsNullOrEmpty(dataCells.FirstOrDefault().Color))
                range.Font.Color = System.Drawing.ColorTranslator.FromHtml(dataCells.FirstOrDefault().Color);
            if (!string.IsNullOrEmpty(dataCells.FirstOrDefault().BgColor))
                range.Interior.Color = System.Drawing.ColorTranslator.FromHtml(dataCells.FirstOrDefault().BgColor);
            if (dataCells.FirstOrDefault().IsCentered)
                range.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;
            if (dataCells.FirstOrDefault().FontSize != 14)
                range.Font.Size = dataCells.FirstOrDefault().FontSize;
            range.Font.Bold = dataCells.FirstOrDefault().IsBold;
        }

        public static void DisplayBatchColumn(ExcelVSTO.Worksheet sheet, List<DataCell> dataCells, int StartRow, int column)
        {
            int n = dataCells.Count;
            object[,] values = new object[n, 1];

            for (int i = 0; i < n; i++)
                values[i, 0] = dataCells[i].Value;

            var range = sheet.Range[
                sheet.Cells[StartRow, column],
                sheet.Cells[StartRow + n - 1, column]
            ];

            range.Value2 = values;
            if (!string.IsNullOrEmpty(dataCells.FirstOrDefault().Color))
                range.Font.Color = System.Drawing.ColorTranslator.FromHtml(dataCells.FirstOrDefault().Color);
            if (!string.IsNullOrEmpty(dataCells.FirstOrDefault().BgColor))
                range.Interior.Color = System.Drawing.ColorTranslator.FromHtml(dataCells.FirstOrDefault().BgColor);
            if (dataCells.FirstOrDefault().IsCentered)
                range.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;
            if (dataCells.FirstOrDefault().FontSize != 14)
                range.Font.Size = dataCells.FirstOrDefault().FontSize;
            range.Font.Bold = dataCells.FirstOrDefault().IsBold;
        }

        public static void DisplayCell(ExcelVSTO.Worksheet sheet, DataCell dataCell, int row, int column)
        {
            var cell = sheet.Cells[row, column] as ExcelInterop.Range;
            cell.Value2 = dataCell.Value;
            if (!string.IsNullOrEmpty(dataCell.Color))
                cell.Font.Color = System.Drawing.ColorTranslator.FromHtml(dataCell.Color);
            if (!string.IsNullOrEmpty(dataCell.BgColor))
                cell.Interior.Color = System.Drawing.ColorTranslator.FromHtml(dataCell.BgColor);
            if (dataCell.IsCentered)
                cell.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;
            if (dataCell.FontSize != 14)
                cell.Font.Size = dataCell.FontSize;
            cell.Font.Bold = dataCell.IsBold;
        }
    }

    public class ColumnData
    {
        public int StartRow { get; set; }
        public int Column { get; set; }
        public List<DataCell> Data { get; set; }

        public ColumnData(int StartRow, int Column, List<DataCell> Data)
        {
            this.StartRow = StartRow;
            this.Column = Column;
            this.Data = Data;
        }
    }

    public class RowData
    {
        public int Row { get; set; }
        public int StartColumn { get; set; }
        public List<DataCell> Data { get; set; }

        public RowData(int Row, int StartColumn, List<DataCell> Data)
        {
            this.Row = Row;
            this.StartColumn = StartColumn;
            this.Data = Data;
        }
    }

    public class BlockData
    {
        public int StartRow { get; set; }
        public int StartColumn { get; set; }
        public List<string> Rows { get; set; }
        public List<string> Columns { get; set; }
        public object[,] Data { get; set; }

        public BlockData(int StartRow, int StartColumn, List<string> Rows, List<string> Columns)
        {
            this.StartRow = StartRow;
            this.StartColumn = StartColumn;
            this.Rows = Rows;
            this.Columns = Columns;
            this.Data = BuildDisplayMatrix();
        }

        private object[,] BuildDisplayMatrix()
        {
            int rows = Rows.Count;
            int cols = Columns.Count;

            object[,] matrix = new object[rows, cols];

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    matrix[r, c] = "-";
                }
            }

            return matrix;
        }

        public void UpdateMatrix(string rowName, string ColumnName, object value)
        {
            int r = Rows.IndexOf(rowName);
            int c = Columns.IndexOf(ColumnName);

            if (r >= 0 && c >= 0)
                Data[r, c] = value;
        }

        public bool HasValue(string rowName, string ColumnName)
        {
            int r = Rows.IndexOf(rowName);
            int c = Columns.IndexOf(ColumnName);

            return Data[r, c].ToString() != "-";
        }
    }

    public class DataCell
    {
        public string Value { get; set; }
        public string BgColor { get; set; }
        public bool IsBold { get; set; }
        public bool IsCentered { get; set; }
        public string Color { get; set; }
        public int Offset { get; set; }
        public int FontSize { get; set; }

        public DataCell(string Value, string Color = "", string BgColor = "", bool IsBold = false, bool IsCentered = false, int Offset = 0, int FontSize = 14)
        {
            this.Value = Value;
            this.Color = Color;
            this.BgColor = BgColor;
            this.IsBold = IsBold;
            this.IsCentered = IsCentered;
            this.Offset = Offset;
            this.FontSize = FontSize;
        }
    }

    public class CellMerge
    {
        public int StartRow { get; set; }
        public int EndRow { get; set; }
        public int StartColumn { get; set; }
        public int EndColumn { get; set; }
        public CellMerge(int StartRow, int EndRow, int StartColumn, int EndColumn)
        {
            this.StartRow = StartRow;
            this.EndRow = EndRow;
            this.StartColumn = StartColumn;
            this.EndColumn = EndColumn;
        }

        public void Run(ExcelVSTO.Worksheet sheet)
        {
            var range = sheet.Range[sheet.Cells[StartRow, StartColumn], sheet.Cells[EndRow, EndColumn]] as ExcelInterop.Range;
            range.Merge();
        }
    }
}
