using System.Collections.Generic;

namespace ExcelUtility
{
    public interface IRow
    {
        int Index { get; }
        double Height { get; set; }

        ICell GetCell(string columnName);
        ICell GetCell(int columnIndex);
        IEnumerable<ICell> GetCells();
    }
}
