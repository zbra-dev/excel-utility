using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelUtility.Utils;

namespace ExcelUtility.Impl
{
    internal class Row : IRow
    {
        public static Row FromExisting(XElementData data, double defaultHeight, SharedStrings sharedStrings)
        {
            return new Row(data, defaultHeight, sharedStrings);
        }

        public static Row New(XElementData data, int index, SharedStrings sharedStrings)
        {
            return new Row(data, index, sharedStrings);
        }

        private List<ICell> cells = null; // lazy load
        private double defaultHeight;
        public XElementData Data { get; private set; }
        private SharedStrings sharedStrings;

        public int Index { get; private set; }
        
        public double Height
        { 
            get 
            {
                var ht = Data["ht"];
                return ht == null ? defaultHeight : double.Parse(ht, NumberFormatInfo.InvariantInfo);
            } 
            set 
            {
                Data.SetAttributeValue("ht", value == defaultHeight ? null : (object)value); 
            } 
        }

        // existing rows constructor
        private Row(XElementData data, double defaultHeight, SharedStrings sharedStrings)
        {
            this.Data = data;
            this.defaultHeight = defaultHeight;
            this.sharedStrings = sharedStrings;
            Index = int.Parse(data["r"], NumberFormatInfo.InvariantInfo);
            data.RemoveAttribute("spans"); // clear spans attribute, will be recalculated
        }

        // new rows constructor
        private Row(XElementData data, int index, SharedStrings sharedStrings)
        {
            this.Data = data;
            this.sharedStrings = sharedStrings;
            Index = index;
            data.SetAttributeValue("r", index);
            data.SetAttributeValue("x14ac", "dyDescent", 0.25); // office 2010 specific attribute
        }

        public IEnumerable<ICell> GetCells()
        {
            LazyLoadCells();
            return cells;
        }

        public ICell GetCell(string columnName)
        {
            LazyLoadCells();
            var search = new FakeCell() { Name = columnName + Index };
            int insert = cells.BinarySearch(search, CompareCells);
            if (insert < 0)
            {
                insert = ~insert;
                XElementData cellData;
                if (insert == 0)
                    cellData = Data.Add("c");
                else
                    cellData = ((Cell)cells[insert - 1]).Data.AddAfterSelf("c");
                cells.Insert(insert, Cell.New(cellData, search.Name, sharedStrings));
            }
            return cells[insert];
        }

        public ICell GetCell(int columnIndex)
        {
            return GetCell(ColumnUtil.GetColumnName(columnIndex));
        }

        private int CompareCells(ICell cell1, ICell cell2)
        {
            int compare = cell1.Name.Length.CompareTo(cell2.Name.Length);
            if (compare == 0)
                return cell1.Name.CompareTo(cell2.Name);
            return compare;
        }

        private void LazyLoadCells()
        {
            if (cells == null)
                cells = Data.Descendants("c").Select(c => (ICell)(Cell.FromExisting(c, sharedStrings))).ToList();
        }

        private class FakeCell : ICell
        {
            public string StringValue { get; set; }
            public double DoubleValue { get; set; }
            public string Name { get; set; }
            public long LongValue { get; set; }
            public bool IsTypeString { get; set; }
            public XElementData Data { get; private set; }
        }

    }
}
