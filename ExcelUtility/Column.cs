using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelUtility.Impl;
using System.Xml.Linq;

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
                index = index + c - 65;
            return index + (26 * (columnReference.Length - 1));
        }

        private void SetWidth(double value)
        {
            // change worksheet XML value
            width = value;
        }
    }
}
