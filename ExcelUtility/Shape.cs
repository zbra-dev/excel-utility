using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using ExcelUtility.Impl;
using System.Xml.Linq;

namespace ExcelUtility
{
    public class Shape
    {
        private XElement shapeData;
        private Drawing drawingContent;
        private XNamespace NamespaceWithPrefixXdr { get { return shapeData.GetNamespaceOfPrefix("xdr"); } }
        private XNamespace NamespaceWithPrefixA { get { return shapeData.GetNamespaceOfPrefix("a"); } }
        private string text;

        public Color ForeColor { get; set; }
        public int MarginLeft { get; set; }
        public int MarginRight { get; set; }
        public int MarginTop { get; set; }
        public int MarginBottom { get; set; }

        public string Text { get { return text; } set { SetText(value); } }

        public Shape(XElement shapeData, Drawing drawingContent)
        {
            this.shapeData = shapeData;
            this.drawingContent = drawingContent;
        }

        private void SetText(string value)
        {
            text = value;
            var textData = shapeData.Descendants(NamespaceWithPrefixA + "t").First();
            textData.SetValue(value);
        }
    }
}
