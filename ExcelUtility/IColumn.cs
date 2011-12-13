
namespace ExcelUtility
{
    public interface IColumn
    {
        string Name { get; }
        long Index { get; }
        double Width { get; set; }
        int? Style { get; set; }
        void Remove();
    }
}
