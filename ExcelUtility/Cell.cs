using System.Xml.Linq;
using System.Linq;

namespace ExcelUtility
{
    public class Cell
    {
        private IWorksheet parentWorksheet;
        private XElement cellData;
        private XNamespace Namespace { get { return cellData.GetDefaultNamespace(); } }

        public string Value { get { return cellData.Descendants(Namespace + "v").First().Value; } }
        public string Name { get; private set; }

        internal Cell(XElement cellData, IWorksheet parentWorksheet, string name)
        {
            this.cellData = cellData;
            this.parentWorksheet = parentWorksheet;
            Name = name;
        }

        public void SetValue(string value)
        {
            if (cellData.Attribute("t") == null || cellData.Attribute("t").Value != "s")
                cellData.SetAttributeValue("t", "s");
            cellData.SetElementValue(Namespace + "v", parentWorksheet.SharedStrings.GetStringReferenceOf(value));
        }

        public void SetValue(double value)
        {
            if (cellData.Attribute("t") != null && cellData.Attribute("t").Value == "s")
                cellData.SetAttributeValue("t", string.Empty);
            cellData.SetElementValue(cellData.GetDefaultNamespace() + "v", value);
        }
    }
}
