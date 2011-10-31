using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelUtility.Impl;

namespace ExcelUtility
{
    public class Row
    {
        private Worksheet worksheet;
        private double height;

        public string Name { get; private set; }
        public double Height { get { return height; } set { SetHeight(value); } }

        internal Row(Worksheet worksheet, string name)
        {
            this.worksheet = worksheet;
            Name = name;
        }

        private void SetHeight(double value)
        {
            // change worksheet XML value
            height = value;
        }
    }
}
