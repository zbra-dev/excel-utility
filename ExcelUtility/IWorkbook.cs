﻿using System.Collections.Generic;
using ExcelUtility.Impl;

namespace ExcelUtility
{
    internal interface IWorkbook
    {
        SharedStrings SharedStrings { get; }
        IEnumerable<IWorksheet> Worksheets { get; }
    }
}
