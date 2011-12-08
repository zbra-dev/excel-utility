using ExcelUtility.Utils;

namespace ExcelUtility.Impl
{
    internal class ShapeProperties
    {
        public static ShapeProperties New(XElementData data, DrawPosition from, DrawPosition to)
        {
            return new ShapeProperties(data, from, to);    
        }

        public XElementData data { get; set; }

        public void SetSolidFill(int r, int g, int b)
        {
            RgbColorModel.SetHexColor(data.Element("a", "solidFill"), r, g, b);
        }

        private ShapeProperties(XElementData data, DrawPosition from, DrawPosition to)
        {
            this.data = data;
            WriteContents(from, to);
        }

        private void WriteContents(DrawPosition from, DrawPosition to)
        {
            // transform
            //var transform = data.Add("a", "xfrm");
            //var offset = transform.Add("off");
            //offset.SetAttributeValue("x", from.X); // absolute coordinate
            //offset.SetAttributeValue("y", from.Y); // absolute coordinate
            //var extension = transform.Add("ext");
            //extension.SetAttributeValue("cx", to.X - from.X); // absolute width
            //extension.SetAttributeValue("cy", to.Y - from.Y); // absolute height
            // preset geometry
            var presetGeometry = data.Add("a", "prstGeom");
            presetGeometry["prst"] = "rect";
            presetGeometry.Add("avLst");
            // solid fill
            RgbColorModel.SetHexColor(data.Add("a", "solidFill"), 255, 255, 255);//TODO: set color
            // outline
            var outline = data.Add("a", "ln");
            outline["w"] = "9525"; // line width
            outline["cmpd"] = "sng"; // compound type (default: sng)
            RgbColorModel.SetHexColor(outline.Add("solidFill"), 0, 0, 0);
        }

    }
}
