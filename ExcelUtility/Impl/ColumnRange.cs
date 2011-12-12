using System.Globalization;
using ExcelUtility.Utils;

namespace ExcelUtility.Impl
{
    internal class ColumnRange
    {
        private XElementData data;

        public long Min { get { return long.Parse(data["min"], NumberFormatInfo.InvariantInfo); } set { data.SetAttributeValue("min", value); } }
        public long Max { get { return long.Parse(data["max"], NumberFormatInfo.InvariantInfo); } set { data.SetAttributeValue("max", value); } }
        private bool CustomWidth { get { return object.Equals(data["customWidth"], "1"); } set { data.SetAttributeValue("customWidth", value ? "1" : null); } }

        public double Width
        {
            get
            {
                var width = data["width"];
                return width == null ? Column.DefaultWidth : double.Parse(width, NumberFormatInfo.InvariantInfo);
            }
            set
            {
                var customWidth = value != Column.DefaultWidth;
                CustomWidth = customWidth;
                data.SetAttributeValue("width", customWidth ? (object)value : null);
            }
        }
        
        public int? Style 
        { 
            get 
            { 
                var value = data["style"];
                if (value == null)
                    return null;
                int style;
                if (int.TryParse(value, out style))
                    return style;
                return null;
            } 
            set 
            { 
                data["style"] = value == null ? null : value.ToString(); 
            } 
        }

        public ColumnRange(XElementData data)
        {
            this.data = data;
        }

        public ColumnRange(XElementData data, long min, long max, double width, int? style)
        {
            this.data = data;
            Min = min;
            Max = max;
            Width = ((int)(width * 256)) / (double)256;
            Style = style;
        }
    }
}
