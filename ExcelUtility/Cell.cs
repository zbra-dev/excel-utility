using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelUtility.Impl;

namespace ExcelUtility
{
    public class Cell
    {
        private Worksheet worksheet;
        private string value;

        public string Value { get { return value; } set { SetValue(value); } }

        public string Name { get; private set; }

        internal Cell(Worksheet worksheet, string name)
        {
            this.worksheet = worksheet;
            Name = name;
        }

        private void SetValue(string value)
        {
            // set value in worksheet XML
            this.value = value;
        }
    }
}
