using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelUtility.Impl;
using System.Xml.Linq;
using System.Globalization;

namespace ExcelUtility
{
    public class Column
    {
        private IWorksheet worksheet;
        private XElement columnData;
        private double width;

        public string Name { get; private set; }
        public int ColumnIndex { get { return GetColumnIndex(Name); } }
        public double Width { get { return width; } set { SetWidth(value); } }
        public double Position { get { return worksheet.GetColumnPosition(ColumnIndex); } }

        internal Column(XElement columnData, IWorksheet worksheet, string name)
        {
            this.columnData = columnData;
            this.worksheet = worksheet;
            Name = name;
        }

        private int GetColumnIndex(string columnReference)
        {
            int index = 0;
            foreach (char c in columnReference)
                index = index + c - 64;
            return index + (26 * (columnReference.Length - 1));
        }

        private void SetWidth(double value)
        {
            width = value;
            int min = Convert.ToInt32(columnData.Attribute("min").Value);
            int max = Convert.ToInt32(columnData.Attribute("max").Value);

            if (min != max)
            {
                double currentWidth = Convert.ToDouble(columnData.Attribute("width").Value, CultureInfo.InvariantCulture);
                columnData.SetAttributeValue("min", ColumnIndex);
                columnData.SetAttributeValue("max", ColumnIndex);
                if (min == ColumnIndex)
                {
                    var col = worksheet.CreateColumnBetweenWith(min + 1, max, currentWidth);
                }
                else if (max == ColumnIndex)
                    worksheet.CreateColumnBetweenWith(min, max - 1, currentWidth);
                else
                {
                    worksheet.CreateColumnBetweenWith(min, ColumnIndex - 1, currentWidth);
                    worksheet.CreateColumnBetweenWith(ColumnIndex + 1, max, currentWidth);
                }
            }
            columnData.SetAttributeValue("width", value);
        }
    }
}
