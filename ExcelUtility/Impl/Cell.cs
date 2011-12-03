using System.Globalization;
using ExcelUtility.Utils;

namespace ExcelUtility.Impl
{
    internal class Cell : ICell
    {
        public static Cell FromExisting(XElementData data, SharedStrings sharedStrings)
        {
            return new Cell(data, sharedStrings);
        }

        public static Cell New(XElementData data, string name, SharedStrings sharedStrings)
        {
            return new Cell(data, name, sharedStrings);
        }

        private SharedStrings sharedStrings;
        private string Type { get { return Data["t"]; } set { Data["t"] = value; } }

        public XElementData Data { get; private set; }
        public bool IsTypeString { get { var type = Type; return type != null && type == "s"; } }
        public string StringValue { get { return GetStringValue(); } set { SetStringValue(value); } }
        public double DoubleValue { get { return GetDoubleValue(); } set { SetDoubleValue(value); } }
        public long LongValue { get { return GetLongValue(); } set { SetLongValue(value); } }
        public string InternalValue { get { return Data.Element("v").Value; } set { Data.SetElementValue("v", value); } }
        public string Name { get; private set; }

        // existing cells constructor
        private Cell(XElementData data, SharedStrings sharedStrings)
        {
            this.Data = data;
            this.sharedStrings = sharedStrings;
            Name = data["r"];
        }

        // new cell constructor
        private Cell(XElementData data, string name, SharedStrings sharedStrings)
        {
            this.Data = data;
            this.sharedStrings = sharedStrings;
            Name = name;
            data["r"] = name;
        }

        private string GetStringValue()
        {
            return InternalValue;
        }

        private void SetStringValue(string value)
        {
            Type = "s";
            InternalValue = sharedStrings.GetStringReferenceOf(value).ToString();
        }

        private double GetDoubleValue()
        {
            return double.Parse(Data.Element("v").Value, NumberFormatInfo.InvariantInfo);
        }

        private void SetDoubleValue(double value)
        {
            Type = null;
            Data.SetElementValue("v", value);
        }

        private long GetLongValue()
        {
            return long.Parse(Data.Element("v").Value, NumberFormatInfo.InvariantInfo);
        }

        private void SetLongValue(long value)
        {
            Type = null;
            Data.SetElementValue("v", value);
        }

        public override string ToString()
        {
            return string.Format("{0}={1}", Name, StringValue);
        }
    }
}
