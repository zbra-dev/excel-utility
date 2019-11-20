using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ExcelUtility.Utils;

namespace ExcelUtility.Impl
{
    internal class Worksheet : IWorksheet, IWorksheetData
    {
        private const int EmuFactor = 12700;

        private XElementData data;
        private XElementData relationshipsData;
        private readonly int sheetId;
        private readonly string worksheetFolder;
        private bool canDraw;

        private Drawings drawings;
        private SheetData sheetData;

        public SheetColumns SheetColumns { get; private set; }
        public IWorkbook Workbook { get; private set; }
        public string Name { get; private set; }
        public double DefaultRowHeight { get; private set; }
        public IEnumerable<IRow> DefinedRows { get { return sheetData.DefinedRows; } }
        public IEnumerable<IColumn> DefinedColumns { get { return SheetColumns.DefinedColumns; } }
        public IEnumerable<IShape> Shapes { get { return drawings.Shapes; } }
        public IEnumerable<ICell> DefinedCells { get { return sheetData.DefinedRows.SelectMany(r => r.DefinedCells); } }
        public ISheetViews SheetView { get; private set; }

        public Worksheet(XElementData data, IWorkbook workbook, string worksheetFolder, string name, int sheetId)
        {
            this.data = data;
            this.Workbook = workbook;
            this.relationshipsData = new XElementData(XDocument.Load(string.Format("{0}/_rels/sheet{1}.xml.rels", worksheetFolder, sheetId)).Root);
            this.worksheetFolder = worksheetFolder;
            this.sheetId = sheetId;
            Name = name;
            ReadContents();
            var dimension = data.Element("dimension");
            if (dimension != null)
                dimension.Remove(); // clear dimension attribute - will be recalculated
        }

        private void ReadContents()
        {
            DefaultRowHeight = double.Parse(data.Element("sheetFormatPr")["defaultRowHeight"], NumberFormatInfo.InvariantInfo);
            var cols = data.Element("cols") ?? data.Element("sheetFormatPr").AddAfterSelf("cols");
            SheetColumns = new SheetColumns(cols);
            sheetData = new SheetData(data.Element("sheetData"), this, SheetColumns);
            SheetView = new SheetViews(data.Element("sheetViews"));
            canDraw = TryLoadDrawings();
        }

        private bool TryLoadDrawings()
        {
            if (data.Element("drawing") == null)
            {
                return false;
            }

            var drawingsId = data.Element("drawing").AttributeValue("r", "id");
            var targetPath = relationshipsData.Descendants("Relationship").Single(r => r["Id"] == drawingsId)["Target"];
            drawings = new Drawings(string.Format("{0}/{1}", worksheetFolder, targetPath));
            return true;
        }

        public IColumn GetColumn(string name)
        {
            return SheetColumns.GetColumn(name);
        }

        public IColumn GetColumn(int index)
        {
            return SheetColumns.GetColumn(index);
        }

        public IRow GetRow(int index)
        {
            return sheetData.GetRow(index);
        }

        public ICell GetCell(string name)
        {
            var match = Regex.Match(name, @"([|A-Z|a-z|]*)([\d]*)");
            if (!match.Groups[1].Success || !match.Groups[2].Success)
                throw new ArgumentException(string.Format("Invalid cell [{0}]", name));
            return sheetData.GetRow(int.Parse(match.Groups[2].Value)).GetCell(match.Groups[1].Value);
        }


        public IShape DrawShape(int columnFrom, double columnFromOffset, int rowFrom, double rowFromOffset, int columnTo, double columnToOffset, int rowTo, double rowToOffset)
        {
            if (!canDraw)
            {
                return null;
            }
            var from = CalculatePosition(columnFrom, columnFromOffset, rowFrom, rowFromOffset);
            var to = CalculatePosition(columnTo, columnToOffset, rowTo, rowToOffset);
            return drawings.DrawShape(from, to);
        }

        private DrawPosition CalculatePosition(int columnIndex, double columnOffset, int rowIndex, double rowOffset)
        {
            var pos = new DrawPosition()
            {
                ColumnIndex = columnIndex,
                ColumnOffset = (int)(columnOffset * EmuFactor),
                RowIndex = rowIndex,
                RowOffset = (int)(rowOffset * EmuFactor),
            };
            //pos.X = ((int)(sheetColumns.GetXPosition(pos.ColumnIndex) * EmuFactor)) + pos.ColumnOffset;
            //pos.Y = ((int)(sheetData.GetYPosition(pos.RowIndex) * EmuFactor)) + pos.RowOffset;
            return pos;
        }

        public void Save()
        {
            if (drawings != null)
                drawings.Save();
            SheetColumns.Save();
            data.Save(string.Format("{0}/sheet{1}.xml", worksheetFolder, sheetId));
        }
    }
}