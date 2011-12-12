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
