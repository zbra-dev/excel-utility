using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelUtility.Impl;

namespace ExcelUtility
{
    public class Column
    {
        private Worksheet worksheet;
        private double width;

        public string Name { get; private set; }
        public double Width { get { return width; } set { SetWidth(value); } }

        internal Column(Worksheet worksheet, string name)
        {
            this.worksheet = worksheet;
            Name = name;
        }

        private void SetWidth(double value)
        {
            // change worksheet XML value
            width = value;
        }
    }
}
