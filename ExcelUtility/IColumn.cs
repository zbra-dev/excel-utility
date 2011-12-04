
namespace ExcelUtility
{
    public interface IColumn
    {
        string Name { get; }
        long Index { get; }
        double Width { get; set; }
        int? InternalColor { get; set; }
        int? Style { get; set; }
    }
}
