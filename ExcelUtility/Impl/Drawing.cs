using System;
using System.Linq;
using System.Xml.Linq;
using System.Drawing;

namespace ExcelUtility.Impl
{
    public class Drawing
    {
        private string pathUrl;
        private XElement drawingData;
        private XNamespace NamespaceWithPrefixXdr { get { return drawingData.GetNamespaceOfPrefix("xdr"); } }
        private XNamespace NamespaceWithPrefixA { get { return drawingData.GetNamespaceOfPrefix("a"); } }
        private IWorksheet parentWorksheet;
        
        public Drawing(string drawingUrl, IWorksheet worksheet)
        {
            pathUrl = drawingUrl;
            drawingData = XElement.Load(drawingUrl);
            parentWorksheet = worksheet;
        }

        public Shape DrawShape(Column column, Row row, double rowOffSet, double colOffSet, double width, double height)
        {
            Column targetColumn = parentWorksheet.CalculateColumnAfter(column, colOffSet, width);
            Row targetRow = parentWorksheet.CalculateRowAfter(row, rowOffSet, height);

            #region From

            var from = new XElement(NamespaceWithPrefixXdr + "from",
                                    new XElement(NamespaceWithPrefixXdr + "col", column.ColumnIndex),
                                    new XElement(NamespaceWithPrefixXdr + "colOff", (int)(colOffSet * 12000)),
                                    new XElement(NamespaceWithPrefixXdr + "row", row.Index),
                                    new XElement(NamespaceWithPrefixXdr + "rowOff", (int)(rowOffSet * 12000)));
            #endregion

            #region To

            var to = new XElement(NamespaceWithPrefixXdr + "to",
                                    new XElement(NamespaceWithPrefixXdr + "col", targetColumn.ColumnIndex),
                                    new XElement(NamespaceWithPrefixXdr + "colOff", (int)((column.Position + colOffSet + width - targetColumn.Position) * 12000)),
                                    new XElement(NamespaceWithPrefixXdr + "row", targetRow.Index),
                                    new XElement(NamespaceWithPrefixXdr + "rowOff", (int)((row.Position + rowOffSet + height - targetRow.Position) * 12000)));
            #endregion

            #region Sp

            #region nvSpPr
            int id = (from d in drawingData.Descendants(NamespaceWithPrefixXdr + "cNvPr")
                      orderby Convert.ToInt32(d.Attribute("id").Value) descending
                      select Convert.ToInt32(d.Attribute("id").Value)).FirstOrDefault() + 1;

            var nonVisualShapeProperties = new XElement(NamespaceWithPrefixXdr + "nvSpPr",
                    new XElement(NamespaceWithPrefixXdr + "cNvPr", new XAttribute("id", id), new XAttribute("name", string.Format("TextBox {0}", id))),
                    new XElement(NamespaceWithPrefixXdr + "cNvSpPr", new XAttribute("txBox", "1")));
            #endregion

            #region spPr

            //TODO: understand the offset calculation
            var graphicalFrame2D = new XElement(NamespaceWithPrefixA + "xfrm",
                    new XElement(NamespaceWithPrefixA + "off", new XAttribute("x", 0), new XAttribute("y", 0)),
                    new XElement(NamespaceWithPrefixA + "ext", new XAttribute("cx", width * 12000), new XAttribute("cy", height * 12000)));

            //TODO: fill w and cmpd values.
            var line = new XElement(NamespaceWithPrefixA + "ln", new XAttribute("w", "9525"), new XAttribute("cmpd", "sng"),//sng=singleLine, dbl=doubleLibe -- 9525 default value to.
                    new XElement(NamespaceWithPrefixA + "solidFill",
                        new XElement(NamespaceWithPrefixA + "srgbClr", new XAttribute("val", String.Format( "{0:x2}{1:x2}{2:x2}", Color.Black.R, Color.Black.G, Color.Black.B )))));//TODO: Fix color scheme

            var shapeProperties = new XElement(NamespaceWithPrefixXdr + "spPr", graphicalFrame2D,
                    new XElement(NamespaceWithPrefixA + "prstGeom", new XAttribute("prst", "rect")),
                    new XElement(NamespaceWithPrefixA + "solidFill", new XElement(NamespaceWithPrefixA + "srgbClr", new XAttribute("val", String.Format("{0:x2}{1:x2}{2:x2}", Color.White.R, Color.White.G, Color.White.B)))), //TODO: Fix color scheme
                    line);

            #endregion

            #region Style

            var styleRgbColor = new XElement(NamespaceWithPrefixA + "scrgbClr", new XAttribute("r", 0), new XAttribute("g", 0), new XAttribute("b", 0));
            var style = new XElement(NamespaceWithPrefixXdr + "style",
                            new XElement(NamespaceWithPrefixA + "lnRef", new XAttribute("idx", 0), styleRgbColor),
                            new XElement(NamespaceWithPrefixA + "fillRef", new XAttribute("idx", 0), styleRgbColor),
                            new XElement(NamespaceWithPrefixA + "effectRef", new XAttribute("idx", 0), styleRgbColor),
                            new XElement(NamespaceWithPrefixA + "fontRef", new XAttribute("idx", "minor"), new XElement(NamespaceWithPrefixA + "schemeClr", new XAttribute("val", "dk1"))));
            #endregion

            #region txBody
            var bodyProperties = new XElement(NamespaceWithPrefixA + "bodyPr", new XAttribute("vertOverflow", "clip"), new XAttribute("horzOverflow", "clip"), new XAttribute("vert", "horz"),
                            new XAttribute("lIns", 63500), new XAttribute("tIns", 25400), new XAttribute("rIns", 63500), new XAttribute("rtlCol", 0), new XAttribute("anchor", "t"));
            var textCharacaterProperties = new XElement(NamespaceWithPrefixA + "rPr", new XAttribute("lang", "en-US"), new XAttribute("sz", "1000"),
                            new XElement(NamespaceWithPrefixA + "solidFill", new XElement(NamespaceWithPrefixA + "srgbClr", new XAttribute("val", String.Format("{0:x2}{1:x2}{2:x2}", Color.Black.R, Color.Black.G, Color.Black.B)))),
                            new XElement(NamespaceWithPrefixA + "latin", new XAttribute("typeface", "Arial")));
            
            var textParagraphs = new XElement(NamespaceWithPrefixA + "p", new XElement(NamespaceWithPrefixA + "r", textCharacaterProperties, new XElement(NamespaceWithPrefixA + "t")));
            var textBody = new XElement(NamespaceWithPrefixXdr + "txBody", bodyProperties, textParagraphs);
            #endregion

            var sp = new XElement(NamespaceWithPrefixXdr + "sp", new XAttribute("macro", ""), new XAttribute("textlink", ""), nonVisualShapeProperties, shapeProperties, style, textBody);
            #endregion

            var shape = new XElement(NamespaceWithPrefixXdr + "twoCellAnchor", from, to, sp,
                new XElement(NamespaceWithPrefixXdr + "clientData"));

            drawingData.Add(shape);

            return new Shape(shape, this);
        }

        public void Save()
        {
            drawingData.Save(pathUrl);
        }

        private int GetColumnIndex(string columnReference)
        {
            int index = 0;
            foreach (char c in columnReference)
                index = index + c - 65;
            return index + (26 * (columnReference.Length - 1));
        }
    }
}
