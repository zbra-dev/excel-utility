using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelUtility.Impl
{
    internal class Worksheet : IWorksheet
    {
        #region IWorksheet Members

        public Column GetColumn(string name)
        {
            throw new NotImplementedException();
        }

        public Cell GetCell(string name)
        {
            throw new NotImplementedException();
        }

        public Row GetRow(string name)
        {
            throw new NotImplementedException();
        }

        public Shape DrawShape(double x, double y, double width, double height)
        {
            throw new NotImplementedException();
        }

        #endregion
    }
}
