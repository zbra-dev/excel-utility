using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelUtility.Impl;

namespace ExcelUtility
{
    internal interface IWorkbook
    {
        SharedStrings SharedStrings { get; }
        IEnumerable<IWorksheet> Worksheets { get; }
    }
}
