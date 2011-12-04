namespace ExcelUtility
{
    public interface ICell
    {
        string StringValue { get; set; }
        double DoubleValue { get; set; }
        long LongValue { get; set; }
        string Name { get; }
        bool IsTypeString { get; }
        int? Style { get; set; }
    }
}
