using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelUtility
{
    public interface IWorksheet
    {
        Column GetColumn(string name);
        Row GetRow(string name);
        Cell GetCell(string name);
        Shape DrawShape(double x, double y, double width, double height);
    }
}
