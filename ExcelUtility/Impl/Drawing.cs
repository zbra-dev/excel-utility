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
            Column destColumn = parentWorksheet.CalculateColumnAfter(column, colOffSet, width);
            Row destRow = parentWorksheet.CalculateRowAfter(row, rowOffSet, height);

            #region From

            var from = new XElement(NamespaceWithPrefixXdr + "from",
                                    new XElement(NamespaceWithPrefixXdr + "col", column.ColumnIndex),
                                    new XElement(NamespaceWithPrefixXdr + "colOff", colOffSet),
                                    new XElement(NamespaceWithPrefixXdr + "row", row),
                                    new XElement(NamespaceWithPrefixXdr + "rowOff", rowOffSet));
            #endregion

            #region To

            var to = new XElement(NamespaceWithPrefixXdr + "to",
                                    new XElement(NamespaceWithPrefixXdr + "col", destColumn.Name),
                                    new XElement(NamespaceWithPrefixXdr + "colOff", column.Position + colOffSet + width - destColumn.Position),
                                    new XElement(NamespaceWithPrefixXdr + "row", destRow),
                                    new XElement(NamespaceWithPrefixXdr + "rowOff", row.Position + rowOffSet + height - destRow.Position));
            #endregion

            #region Sp

            #region nvSpPr
            int id = (from d in drawingData.Descendants(NamespaceWithPrefixXdr + "cNvPr")
                      orderby Convert.ToInt32(d.Attribute("id").Value) descending
                      select Convert.ToInt32(d.Attribute("id").Value)).FirstOrDefault() + 1;

            var nvSpPr = new XElement(NamespaceWithPrefixXdr + "nvSpPr",
                    new XElement(NamespaceWithPrefixXdr + "cNvPr", new XAttribute("id", id), new XAttribute("name", string.Format("TextBox {0}", id))),
                    new XElement(NamespaceWithPrefixXdr + "cNvSpPr", new XAttribute("txBox", "1")));
            #endregion

            #region spPr

            //TODO: understand the offset calculation
            var xfrm = new XElement(NamespaceWithPrefixA + "xfrm",
                    new XElement(NamespaceWithPrefixA + "off", new XAttribute("x", 0), new XAttribute("y", 0)),
                    new XElement(NamespaceWithPrefixA + "ext", new XAttribute("cx", width * 12000), new XAttribute("cy", height * 12000)));

            //TODO: fill w and cmpd values.
            var ln = new XElement(NamespaceWithPrefixA + "ln", new XAttribute("w", ""), new XAttribute("cmpd", ""),
                    new XElement(NamespaceWithPrefixA + "solidFill",
                        new XElement(NamespaceWithPrefixA + "srgbClr", new XAttribute("val", Color.Black.ToArgb() & 0x00FFFFFF))));

            var spPr = new XElement(NamespaceWithPrefixXdr + "spPr", xfrm,
                    new XElement(NamespaceWithPrefixA + "prstGeom", new XAttribute("prst", "rect")),
                    new XElement(NamespaceWithPrefixA + "solidFill", new XElement(NamespaceWithPrefixA + "srgbClr", new XAttribute("val", Color.White.ToArgb() & 0x00FFFFFF))),
                    ln);

            #endregion

            #region Style

            var styleScrgbClr = new XElement(NamespaceWithPrefixA + "scrgbClr", new XAttribute("r", 0), new XAttribute("g", 0), new XAttribute("b", 0));
            var style = new XElement(NamespaceWithPrefixXdr + "style",
                            new XElement(NamespaceWithPrefixA + "lnRef", new XAttribute("idx", 0), styleScrgbClr),
                            new XElement(NamespaceWithPrefixA + "fillRef", new XAttribute("idx", 0), styleScrgbClr),
                            new XElement(NamespaceWithPrefixA + "effectRef", new XAttribute("idx", 0), styleScrgbClr),
                            new XElement(NamespaceWithPrefixA + "fontRef", new XAttribute("idx", "minor"), new XElement("schemeClr", new XAttribute("val", "dk1"))));
            #endregion

            #region txBody
            var bodyPr = new XElement(NamespaceWithPrefixA + "bodyPr", new XAttribute("vertOverflow", "clip"), new XAttribute("horzOverflow", "clip"), new XAttribute("vert", "horz"),
                            new XAttribute("lIns", 63500), new XAttribute("tIns", 25400), new XAttribute("rIns", 63500), new XAttribute("rtlCol", 0), new XAttribute("anchor", "t"));
            var rPr = new XElement(NamespaceWithPrefixA + "rPr", new XAttribute("lang", "en-US"), new XAttribute("sz", "1000"), 
                            new XElement(NamespaceWithPrefixA + "solidFill", new XElement(NamespaceWithPrefixA + "srgbClr", new XAttribute("val", Color.Black.ToArgb() & 0x00FFFFFF))),
                            new XElement(NamespaceWithPrefixA + "latin", new XAttribute("typeface", "Arial")));
            
            var p = new XElement(NamespaceWithPrefixA + "p", new XElement(NamespaceWithPrefixA + "r", rPr, new XElement(NamespaceWithPrefixA + "t")));
            var txBody = new XElement(NamespaceWithPrefixXdr + "txBody", bodyPr, p);
            #endregion

            var sp = new XElement(NamespaceWithPrefixXdr + "sp", new XAttribute("macro", ""), new XAttribute("textlink", ""), nvSpPr, spPr, style, txBody);
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
