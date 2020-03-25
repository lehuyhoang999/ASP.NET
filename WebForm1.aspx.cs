using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using DevExpress.Spreadsheet;
using DevExpress.XtraSpreadsheet.Export;
using System.IO;

namespace WebApplicationDemo
{
    public class Ship
    {
        public string Description { get; set; }
        public uint Quantity { get; set; }
        public double? Discount { get; set; }
        public double Price { get; set; }
    }
    public partial class WebForm1 : System.Web.UI.Page
    {
        const string DefaultFontName = "Segoe UI";
        const float DefaultFontSize = 10.5f;
        static IList Invoices = new List<Ship>(){
            new Ship() { Description = "Chai", Quantity = 12, Discount = 0.1, Price = 18 },
            new Ship { Description = "Konbu", Quantity = 4, Price = 6 },
            new Ship { Description = "Sir Rodney's Scones", Quantity = 19, Price = 10 },
            new Ship { Description = "Guaraná Fantástica", Quantity = 16, Discount = 0.15, Price = 4.5 },
            new Ship { Description = "Gnocchi di nonna Alice", Quantity = 12, Price = 38 },
            new Ship { Description = "Röd Kaviar", Quantity = 19, Price = 34.80 },
            new Ship { Description = "Konbu", Quantity = 8, Discount = 0.2, Price = 15 },
            new Ship { Description = "Original Frankfurter grüne Soße", Quantity = 2, Price = 13 }
        };
        protected void Page_Load(object sender, EventArgs e)
        {



        }

        protected void btnClick(object sender, EventArgs e)
        {
            IWorkbook book = Spreadsheet.Document;
            book.CreateNewDocument();
            if (book.Worksheets.Count == 1) book.Worksheets.Add();
            book.Worksheets.ActiveWorksheet = book.Worksheets[0];
            book.Unit = DevExpress.Office.DocumentUnit.Point;
            book.Styles.DefaultStyle.Font.Name = DefaultFontName;
            book.Styles.DefaultStyle.Font.Size = DefaultFontSize;
            book.Styles.DefaultStyle.Alignment.Vertical = SpreadsheetVerticalAlignment.Center;

            Worksheet sheet = book.Worksheets[1];
            sheet.ActiveView.ShowGridlines = false;
            sheet.Name = "Invoice";
            PrepareColumns(sheet);
            PrepareWatermarkStyleCell(sheet.Cells[1, 1]);
            FillInvoice(sheet);
            using (FileStream stream = new FileStream("C:\\Users\\LeHuyHoang\\Documents\\SavedDocumentHoangSX.xlsx",
                FileMode.Create, FileAccess.ReadWrite))
            {
                book.SaveDocument(stream, DocumentFormat.Xlsx);
            }
        }
        static void PrepareColumns(Worksheet sheet)
        {
            sheet.Columns[0].WidthInCharacters = 3;
            sheet.Columns[1].WidthInCharacters = 47.86;
            sheet.Columns[2].WidthInCharacters = 12;
            sheet.Columns[3].WidthInCharacters = 18;
            sheet.Columns[4].WidthInCharacters = 16;
            sheet.Columns[5].WidthInCharacters = 21;
        }
        static void PrepareWatermarkStyleCell(Cell cell)
        {
            cell.Font.Size = 36;
            cell.Font.FontStyle = SpreadsheetFontStyle.BoldItalic;
            cell.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Left;
            cell.Alignment.Vertical = SpreadsheetVerticalAlignment.Bottom;
            cell.Font.Color = DevExpress.Utils.DXColor.FromArgb(0xff, 0xE0, 0xE0, 0xE0);
            cell.Value = "INVOICE";
        }

        static void FillInvoice(Worksheet sheet)
        {
            CreateTableColumns(sheet);
            FillInvoiceTable(sheet);
            PrepareTotalValue(sheet);
        }
        const int
            InitialColumnOffset = 1,
            InitialRowOffset = 6;
        static void CreateTableColumns(Worksheet sheet)
        {
            string[] columnTitles = new string[] { "Description", "QTY", "Price", "Discount", "Amount" };
            for (int columnOffset = InitialColumnOffset; columnOffset < InitialColumnOffset + columnTitles.Length; columnOffset++)
            {
                Cell cell = sheet.Cells[InitialRowOffset, columnOffset];
                cell.Font.FontStyle = SpreadsheetFontStyle.Bold;
                cell.Font.Color = DevExpress.Utils.DXColor.FromArgb(0xff, 0x33, 0x33, 0x33);
                cell.Borders.BottomBorder.LineStyle = BorderLineStyle.Medium;
                cell.Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
                cell.Alignment.Horizontal = columnOffset == InitialColumnOffset ? SpreadsheetHorizontalAlignment.Left : SpreadsheetHorizontalAlignment.Center;
                cell.Value = columnTitles[columnOffset - InitialColumnOffset];
            }
        }
        static void FillInvoiceTable(Worksheet sheet)
        {
            int currentRowOffset = InitialRowOffset + 1;
            foreach (Ship ship in Invoices)
            {
                sheet[currentRowOffset, InitialColumnOffset].Value = ship.Description;

                Cell cell = sheet[currentRowOffset, InitialColumnOffset + 1];
                cell.Value = ship.Quantity;
                cell.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;

                cell = sheet[currentRowOffset, InitialColumnOffset + 2];
                sheet[currentRowOffset, InitialColumnOffset + 2].Value = ship.Price;
                cell.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;

                cell = sheet[currentRowOffset, InitialColumnOffset + 3];
                cell.NumberFormat = "0.00%;[Red]-0.00%;;@";
                cell.Value = CellValue.TryCreateFromObject(ship.Discount);
                cell.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;

                cell = sheet[currentRowOffset, InitialColumnOffset + 4];
                cell.Formula = string.Format("C{0}*D{0}*(1-E{0})", currentRowOffset + 1);
                cell.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;

                if (currentRowOffset % 2 == 0)
                    sheet.Range.FromLTRB(1, currentRowOffset, 5, currentRowOffset).FillColor = System.Drawing.Color.FromArgb(0xff, 0xF1, 0xF1, 0xF1);
                currentRowOffset++;
            }
        }
        static void PrepareTotalValue(Worksheet sheet)
        {
            int currentRowOffset = InitialRowOffset + Invoices.Count + 2;
            Cell cell = sheet[currentRowOffset, 4];
            cell.Value = "Total";
            cell.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
            cell.Font.Bold = true;
            cell.Font.Size = 13.5;

            cell = sheet[currentRowOffset, 5];
            cell.Formula = string.Format("SUM(F{0}:F{1})", InitialRowOffset + 2, currentRowOffset - 1);
            cell.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* \" - \"??_);_(@_)";
            cell.Font.Color = System.Drawing.Color.Black;
            cell.Fill.BackgroundColor = System.Drawing.Color.FromArgb(0xff, 0xea, 0xea, 0xea);
            cell.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
            cell.Font.Size = 13.5;
            cell.Select();
        }
    }
}