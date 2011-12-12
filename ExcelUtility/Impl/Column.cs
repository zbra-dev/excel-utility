using ExcelUtility.Utils;

namespace ExcelUtility.Impl
{
    internal class Column : IColumn
    {
        public const double DefaultWidth = 9.140625;

        public string Name { get; private set; }
        public long Index { get; private set; }
        public double Width { get; set; }
        public int? Style { get; set; }

        public Column(string name, long index, double width, int? style)
        {
            Name = name;
            Index = index;
            Width = width;
            Style = style;
        }

        public Column(string name, long index)
            : this(name, index, DefaultWidth, null)
        {
        }

        public ColumnRange ToRange(XElementData data)
        {
            return new ColumnRange(data, Index + 1, Index + 1, Width, Style);
        }
    }
}
