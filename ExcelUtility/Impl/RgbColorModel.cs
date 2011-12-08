using ExcelUtility.Utils;

namespace ExcelUtility.Impl
{
    public class RgbColorModel
    {
        public static void SetHexColor(XElementData data, int r, int g, int b)
        {
            var colorData = data.Element("a", "srgbClr");
            if (colorData == null)
                colorData = data.Add("a", "srgbClr");
            colorData.SetAttributeValue("val", string.Format("{0:X2}{1:X2}{2:X2}", r, g, b));
        }

        public static void SetPercentageColor(XElementData data, int r, int g, int b)
        {
            var colorData = data.Element("a", "scrgbClr");
            if (colorData == null)
                colorData = data.Add("a", "scrgbClr");
            colorData["r"] = r.ToString();
            colorData["g"] = g.ToString();
            colorData["b"] = b.ToString();
        }
    }
}
