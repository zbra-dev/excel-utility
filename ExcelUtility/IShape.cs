
namespace ExcelUtility
{
    public interface IShape
    {
        string Text { get; set; }

        void Remove();
        void SetSolidFill(int r, int g, int b);
    }
}
