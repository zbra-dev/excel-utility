using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelUtility.Impl;

namespace ExcelUtility
{
    internal interface IWorksheetData
    {
        IWorkbook Workbook { get; }
        SheetColumns SheetColumns { get; }
        double DefaultRowHeight { get; }
    }
}
