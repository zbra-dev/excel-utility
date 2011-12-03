using ExcelUtility.Utils;

namespace ExcelUtility.Impl
{
    public class RgbColorModel
    {
        public static void AddHexColor(XElementData data, int r, int g, int b)
        {
            data.Add("a", "srgbClr")["val"] = string.Format("{0:X2}{1:X2}{2:X2}", r, g, b);
        }

        public static void AddPercentageColor(XElementData data, int r, int g, int b)
        {
            var colorData = data.Add("a", "scrgbClr");
            colorData["r"] = r.ToString();
            colorData["g"] = g.ToString();
            colorData["b"] = b.ToString();
        }
    }
}
