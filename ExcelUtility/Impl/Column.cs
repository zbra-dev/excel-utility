using ExcelUtility.Utils;

namespace ExcelUtility.Impl
{
    internal class Column : IColumn
    {
        public const double DefaultWidth = 9.140625;

        public string Name { get; private set; }
        public long Index { get; private set; }
        public double Width { get; set; }
        public int? InternalColor { get; set; }
        public int? Style { get; set; }

        public Column(string name, long index, double width)
        {
            Name = name;
            Index = index;
            Width = width;
            InternalColor = null;
            Style = null;
        }

        public Column(string name, long index)
            : this(name, index, DefaultWidth)
        {
        }

        public ColumnRange ToRange(XElementData data)
        {
            var range = new ColumnRange(data, Index + 1, Index + 1, Width);
            if (Style != null)
                range.Style = Style.ToString();
            if (Width != DefaultWidth)
                range.CustomWidth = true;
            return range;
        }
    }
}
