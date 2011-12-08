using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ExcelUtility.Utils;

namespace ExcelUtility.Impl
{
    internal class Worksheet : IWorksheet
    {
        private const int EmuFactor = 12700;

        private XElementData data;
        private XElementData relationshipsData;
        private double defaultRowHeight;
        private int sheetId;
        private string worksheetFolder;

        private Drawings drawings;
        private SharedStrings sharedStrings;
        private SheetColumns sheetColumns;
        private SheetData sheetData;
        
        public string Name { get; private set; }
        public IEnumerable<IRow> DefinedRows { get { return sheetData.DefinedRows; } }
        public IEnumerable<IColumn> DefinedColumns { get { return sheetColumns.DefinedColumns; } }
        public IEnumerable<IShape> Shapes { get { return drawings.Shapes; } }
        public ISheetViews SheetView { get; private set; }

        public Worksheet(XElementData data, string worksheetFolder, SharedStrings sharedStrings, string name, int sheetId)
        {
            this.data = data;
            this.relationshipsData = new XElementData(XDocument.Load(string.Format("{0}/_rels/sheet{1}.xml.rels", worksheetFolder, sheetId)).Root);
            this.worksheetFolder = worksheetFolder;
            this.sheetId = sheetId;
            this.sharedStrings = sharedStrings;
            Name = name;
            ReadContents();
            var dimension = data.Element("dimension");
            if (dimension != null)
                dimension.Remove(); // clear dimension attribute - will be recalculated
        }

        private void ReadContents()
        {
            defaultRowHeight = double.Parse(data.Element("sheetFormatPr")["defaultRowHeight"], NumberFormatInfo.InvariantInfo);
            var cols = data.Element("cols") ?? data.Element("sheetFormatPr").AddAfterSelf("cols");
            sheetColumns = new SheetColumns(cols);
            sheetData = new SheetData(data.Element("sheetData"), defaultRowHeight, sharedStrings, sheetColumns);
            SheetView = new SheetViews(data.Element("sheetViews"));
            LoadDrawings();
        }

        private void LoadDrawings()
        {
            var drawingsId = data.Element("drawing").AttributeValue("r", "id");
            var targetPath = relationshipsData.Descendants("Relationship").Single(r => r["Id"] == drawingsId)["Target"];
            drawings = new Drawings(string.Format("{0}/{1}", worksheetFolder, targetPath));
        }

        public IColumn GetColumn(string name)
        {
            return sheetColumns.GetColumn(name);
        }

        public IColumn GetColumn(int index)
        {
            return sheetColumns.GetColumn(index);
        }

        public IRow GetRow(int index)
        {
            return sheetData.GetRow(index);
        }

        public ICell GetCell(string name)
        {
            var match = Regex.Match(name, @"(\D)(\d)");
            if (!match.Groups[1].Success || !match.Groups[2].Success)
                throw new ArgumentException(string.Format("Invalid cell [{0}]", name));
            return sheetData.GetRow(int.Parse(match.Groups[2].Value)).GetCell(match.Groups[1].Value);
        }


        public IShape DrawShape(int columnFrom, double columnFromOffset, int rowFrom, double rowFromOffset, int columnTo, double columnToOffset, int rowTo, double rowToOffset)
        {
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
            sheetColumns.Save();
            data.Save(string.Format("{0}/sheet{1}.xml", worksheetFolder, sheetId));
        }

    }
}