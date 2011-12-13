using ExcelUtility.Utils;

namespace ExcelUtility.Impl
{
    internal class Column : IColumn
    {
        public const double DefaultWidth = 9.140625;
        
        private SheetColumns sheetColumns;
        
        public string Name { get; private set; }
        public long Index { get; private set; }
        public double Width { get; set; }
        public int? Style { get; set; }
        
        public Column(string name, long index, double width, int? style, SheetColumns sheetColumns)
        {
            Name = name;
            Index = index;
            Width = width;
            Style = style;
            this.sheetColumns = sheetColumns;
        }

        public Column(string name, long index, SheetColumns sheetColumns)
            : this(name, index, DefaultWidth, null, sheetColumns)
        {
        }

        public ColumnRange ToRange(XElementData data)
        {
            return new ColumnRange(data, Index + 1, Index + 1, Width, Style);
        }

        public void Remove()
        {
            sheetColumns.Remove(this);
        }
    }
}