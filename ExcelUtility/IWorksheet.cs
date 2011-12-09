using System.Collections.Generic;

namespace ExcelUtility
{
    public interface IWorksheet
    {
        string Name { get; }
        double DefaultRowHeight { get; }
        IEnumerable<IRow> DefinedRows { get; }
        IEnumerable<IColumn> DefinedColumns { get; }
        IEnumerable<IShape> Shapes { get; }
        ISheetViews SheetView { get; }
        
        IColumn GetColumn(string name);
        IColumn GetColumn(int index);
        IRow GetRow(int index);
        ICell GetCell(string name);
        IShape DrawShape(int columnFrom, double columnFromOffset, int rowFrom, double rowFromOffset, int columnTo, double columnToOffset, int rowTo, double rowToOffset);

        void Save();
    }
}
