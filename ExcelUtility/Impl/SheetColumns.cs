using System.Collections.Generic;
using System.Linq;
using ExcelUtility.Utils;

namespace ExcelUtility.Impl
{
    internal class SheetColumns
    {
        private XElementData data;
        private List<Column> columns = new List<Column>();

        public IEnumerable<IColumn> DefinedColumns { get { return columns.Cast<IColumn>(); } }

        public SheetColumns(XElementData data)
        {
            this.data = data;
            ReadColumns();
        }

        private void ReadColumns()
        {
            foreach (var columnData in data.Descendants("col"))
            {
                ColumnRange range = new ColumnRange(columnData);
                for (long i = range.Min - 1; i < range.Max; ++i)
                {
                    Column column = new Column(ColumnUtil.GetColumnName(i), i, range.Width, range.Style, this);
                    columns.Add(column);
                }
            }
        }

        public IColumn GetColumn(string name)
        {
            long index = ColumnUtil.GetColumnIndex(name);
            return GetColumn(index, name);
        }

        public IColumn GetColumn(long index)
        {
            string name = ColumnUtil.GetColumnName(index);
            return GetColumn(index, name);
        }

        private Column GetColumn(long index, string name)
        {
            Column newColumn = new Column(name, index, this);
            int insert = columns.BinarySearch(newColumn, CompareColumns);
            if (insert < 0)
            {
                insert = ~insert;
                columns.Insert(insert, newColumn);
            }
            return columns[insert];
        }

        private int CompareColumns(Column column1, Column column2)
        {
            return column1.Index.CompareTo(column2.Index);
        }

        public void Save()
        {
            // if there are no columns remove tag "cols"
            if (columns.Count == 0)
            {
                data.Remove();
                return;
            }

            // clear existing columns
            data.RemoveNodes();

            // recalculate columns range
            ColumnRange lastRange = null;
            Column lastColumn = null;
            foreach (var column in columns)
            {
                if (lastRange == null || !AreEqual(lastColumn, column))
                {
                    lastRange = column.ToRange(data.Add("col"));
                    lastColumn = column;
                }
                else
                {
                    lastRange.Max = column.Index + 1;
                }
            }
        }

        private bool AreEqual(Column c1, Column c2)
        {
            return c1.Width == c2.Width && object.Equals(c1.Style, c2.Style);
        }

        public double GetXPosition(long index)
        {
            var search = new Column("", index, this);
            int insert = columns.BinarySearch(search, CompareColumns);
            if (insert < 0)
                insert = ~insert;
            return columns.Take(insert).Sum(c => c.Width) + ((index - insert) * Column.DefaultWidth);
        }

        public void Remove(IColumn column)
        {
            columns.Remove((Column)column);
        }
    }
}
