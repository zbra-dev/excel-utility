using System.Xml.Linq;

namespace ExcelUtility
{
    public class Cell
    {
        private string value;
        private IWorksheet parentWorksheet;
        private XElement cellData;
        private XNamespace Namespace { get { return cellData.GetDefaultNamespace(); } }

        public string Value { get { return value; } set { SetValue(value); } }
        public string Name { get; private set; }

        internal Cell(XElement cellData, IWorksheet parentWorksheet, string name)
        {
            this.cellData = cellData;
            this.parentWorksheet = parentWorksheet;
            Name = name;
        }

        private void SetValue(string value)
        {
            double numeric;
            if (!double.TryParse(value, out numeric))
            {
                if (cellData.Attribute("t") == null || cellData.Attribute("t").Value != "s")
                    cellData.SetAttributeValue("t", "s");
                cellData.SetElementValue(Namespace + "v", parentWorksheet.SharedStrings.GetStringReferenceOf(value));
            }
            else
            {
                if (cellData.Attribute("t") != null && cellData.Attribute("t").Value == "s")
                    cellData.SetAttributeValue("t", string.Empty);
                cellData.SetElementValue(cellData.GetDefaultNamespace() + "v", value);
            }
            this.value = value;
        }
    }
}
