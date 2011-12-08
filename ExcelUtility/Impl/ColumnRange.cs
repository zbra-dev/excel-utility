using System.Globalization;
using ExcelUtility.Utils;

namespace ExcelUtility.Impl
{
    internal class ColumnRange
    {
        private XElementData data;

        public long Min { get { return long.Parse(data["min"], NumberFormatInfo.InvariantInfo); } set { data.SetAttributeValue("min", value); } }
        public long Max { get { return long.Parse(data["max"], NumberFormatInfo.InvariantInfo); } set { data.SetAttributeValue("max", value); } }
        public double Width { get { return double.Parse(data["width"], NumberFormatInfo.InvariantInfo); } set { data.SetAttributeValue("width", value); } }
        public bool CustomWidth { get { return object.Equals(data["customWidth"], "1"); } set { data["customWidth"] = value ? "1" : null; } }
        public string Style { get { return data["style"]; } set { data["style"] = value; } }

        public ColumnRange(XElementData data)
        {
            this.data = data;
        }

        public ColumnRange(XElementData data, long min, long max, double width)
        {
            this.data = data;
            Min = min;
            Max = max;
            Width = ((int)(width * 256)) / (double)256;
        }
    }
}
