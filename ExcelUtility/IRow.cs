using System.Collections.Generic;

namespace ExcelUtility
{
    public interface IRow
    {
        int Index { get; }
        double Height { get; set; }
        IEnumerable<ICell> DefinedCells { get; }

        ICell GetCell(string columnName);
        ICell GetCell(int columnIndex);
        void Remove(ICell cell);
    }
}
