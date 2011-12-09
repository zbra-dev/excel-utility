using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelUtility
{
    internal interface ISheetData
    {
        IWorksheetData Worksheet { get; }
    }
}
