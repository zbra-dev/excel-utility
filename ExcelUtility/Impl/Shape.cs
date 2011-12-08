using ExcelUtility.Utils;

namespace ExcelUtility.Impl
{
    internal class Shape : IShape
    {
        public static Shape FromExisting(XElementData data)
        {
            return new Shape(data);
        }

        public static Shape New(XElementData data, int id, DrawPosition from, DrawPosition to)
        {
            return new Shape(data, id, from, to);
        }

        private XElementData data;

        public int Id { get { return int.Parse(data.ElementAt("sp.nvSpPr.cNvPr")["id"]); } }
        public string Text { get { return data.ElementAt("sp.txBody.a:p.r.t").Value; } set { data.ElementAt("sp.txBody.a:p.r.t").Value = value; } }
        public ShapeProperties ShapeProperties { get; private set; }

        private Shape(XElementData data)
        {
            this.data = data;
            ShapeProperties = null; // shape properties for existing shapes will not be initialized
        }

        private Shape(XElementData data, int id, DrawPosition from, DrawPosition to)
        {
            this.data = data;
            SetPositions(from, to);
            WriteContents(id, from, to);
        }

        public void SetPositions(DrawPosition from, DrawPosition to)
        {
            SetPosition("from", from);
            SetPosition("to", to);
        }

        private void SetPosition(string name, DrawPosition pos)
        {
            var posData = data.Add(name);
            posData.SetElementValue("col", pos.ColumnIndex);
            posData.SetElementValue("colOff", pos.ColumnOffset);
            posData.SetElementValue("row", pos.RowIndex - 1);
            posData.SetElementValue("rowOff", pos.RowOffset);
        }

        private void WriteContents(int id, DrawPosition from, DrawPosition to)
        {
            var shape = data.Add("sp");
            shape["macro"] = "";
            shape["textlink"] = "";

            // non visual shape properties
            var nonVisualProperties = shape.Add("nvSpPr");
            var drawingProperties = nonVisualProperties.Add("cNvPr");
            drawingProperties.SetAttributeValue("id", id);
            drawingProperties["name"] = "TextBox " + id;
            nonVisualProperties.Add("cNvSpPr")["txBox"] = "1";

            // shape properties
            var shapePropertiesData = shape.Add("spPr");
            ShapeProperties = ShapeProperties.New(shapePropertiesData, from, to);
           
            // style
            var style = shape.Add("style");
            var lineReference = style.Add("a", "lnRef");
            lineReference["idx"] = "0";
            RgbColorModel.AddPercentageColor(lineReference, 0, 0, 0);
            var fillReference = style.Add("a", "fillRef");
            fillReference["idx"] = "0";
            RgbColorModel.AddPercentageColor(fillReference, 0, 0, 0);
            var effectReference = style.Add("a", "effectRef");
            effectReference["idx"] = "0";
            RgbColorModel.AddPercentageColor(effectReference, 0, 0, 0);
            var fontReference = style.Add("a", "fontRef");
            fontReference["idx"] = "minor";
            fontReference.Add("schemeClr")["val"] = "dk1";

            // text body
            var textBody = shape.Add("txBody");
            textBody.Add("a", "bodyPr").SetAttributeValues("vertOverflow=clip horzOverflow=clip vert=horz lIns=63500 tIns=25400 rIns=63500 rtlCol=0 anchor=t");
            textBody.Add("a", "lstStyle");
            var run = textBody.Add("a", "p").Add("r");
            var runProperties = run.Add("rPr");
            runProperties.SetAttributeValues("lang=en-US sz=1000");
            RgbColorModel.AddHexColor(runProperties.Add("solidFill"), 0, 0, 0);
            runProperties.Add("latin")["typeface"] = "Arial";
            run.Add("t").Value = "";

            data.Add("clientData");
        }

        public void Remove()
        {
            data.Remove();
        }

    }
}
