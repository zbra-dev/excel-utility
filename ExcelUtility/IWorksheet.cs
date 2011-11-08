﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using ExcelUtility.Impl;

namespace ExcelUtility
{
    public interface IWorksheet
    {
        SharedStrings SharedStrings { get; set; }
        string Name { get; set; }
        int SheetId { get; set; }
        
        Column GetColumn(string name);
        Row GetRow(string name);
        Cell GetCell(string name);
        Shape DrawShape(double x, double y, double width, double height);
        void SaveChanges(string xmlPath);
    }
}
