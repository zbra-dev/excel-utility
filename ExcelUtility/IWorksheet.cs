using System.Collections.Generic;

namespace ExcelUtility
{
    public interface IWorksheet
    {
        string Name { get; }
        
        IColumn GetColumn(string name);
        IColumn GetColumn(int index);
        IRow GetRow(int index);
        IEnumerable<IRow> GetRows();
        ICell GetCell(string name);
        IShape DrawShape(int columnFrom, double columnFromOffset, int rowFrom, double rowFromOffset, int columnTo, double columnToOffset, int rowTo, double rowToOffset);

        void Save();
    }
}
