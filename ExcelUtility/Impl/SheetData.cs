using System;
using System.Collections.Generic;
using System.Linq;
using ExcelUtility.Utils;

namespace ExcelUtility.Impl
{
    internal class SheetData : ISheetData
    {
        private XElementData data;
        private SheetColumns sheetColumns;
        private List<IRow> rows;

        public IWorksheetData Worksheet { get; private set; }
        public IEnumerable<IRow> DefinedRows { get { return rows; } }

        public SheetData(XElementData data, IWorksheetData worksheet, SheetColumns sheetColumns)
        {
            this.data = data;
            this.Worksheet = worksheet;
            this.sheetColumns = sheetColumns;
            rows = data.Descendants("row").Select(r => ((IRow)Row.FromExisting(r, this))).ToList();
        }

        public IRow GetRow(int index)
        {
            if (index == 0)
                throw new ArgumentException("Row index can't be zero (0)", "index");
            var search = new FakeRow() { Index = index };
            int insert = rows.BinarySearch(search, CompareRows);
            if (insert < 0)
            {
                insert = ~insert;
                XElementData rowData;
                if (insert == 0)
                    rowData = data.Add("row");
                else
                    rowData = ((Row)rows[insert - 1]).Data.AddAfterSelf("row");
                rows.Insert(insert, Row.New(rowData, index, this));
            }
            return rows[insert];
        }

        private int CompareRows(IRow row1, IRow row2)
        {
            return row1.Index.CompareTo(row2.Index);
        }

        public double GetYPosition(int index)
        {
            var search = new FakeRow() { Index = index };
            int insert = rows.BinarySearch(search, CompareRows);
            if (insert < 0)
                insert = ~insert;
            return rows.Take(insert).Sum(r => r.Height) + ((index - insert) * Worksheet.DefaultRowHeight);
        }

        private class FakeRow : IRow
        {
            public int Index { get; set; }
            public double Height { get; set; }
            public IEnumerable<ICell> DefinedCells { get; set; }

            public ICell GetCell(string columnName)
            {
                throw new NotImplementedException();
            }

            public ICell GetCell(int columnIndex)
            {
                throw new NotImplementedException();
            }

            public IEnumerable<ICell> GetCells()
            {
                throw new NotImplementedException();
            }
        }

    }
}
