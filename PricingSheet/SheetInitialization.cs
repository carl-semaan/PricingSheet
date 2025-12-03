using Microsoft.Office.Tools.Excel;
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

        public SheetInitialization() { }

        public SheetInitialization(ExcelVSTO.Worksheet Sheet, string Name, bool ClearOnStartUp, List<SheetButton> SheetButtons)
        {
            this.Sheet = Sheet;
            this.Name = Name;
            this.ClearOnStartUp = ClearOnStartUp;
            this.SheetButtons = SheetButtons;
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
}
