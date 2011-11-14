using System.Collections.Generic;
using ExcelUtility.Impl;

namespace ExcelUtility
{
    public interface IWorksheet
    {
        Drawing Drawing { get; set; }
        SharedStrings SharedStrings { get; set; }
        string Name { get; set; }
        int SheetId { get; set; }
        
        Column GetColumn(string name);
        Column CalculateColumnAfter(Column columnBase, double colOffSet, double width);
        double GetColumnPosition(int columnIndex);

        Row GetRow(int index);
        Row CalculateRowAfter(Row row, double rowOffSet, double height);
        double GetRowPosition(int rowIndex); 
        
        Cell GetCell(string name);
        
        void SaveChanges(string xmlPath);
        void RemoveUnusedStringReferences(IList<StringReference> unusedStringRefences);
        void UpdateStringReferences(IList<StringReference> stringRefences);
        int CountStringsUsed();
    }
}
