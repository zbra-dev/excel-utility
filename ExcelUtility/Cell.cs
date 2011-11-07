using System.Xml.Linq;

namespace ExcelUtility
{
    public class Cell
    {
        private string value;
        
        public XElement XElementCell { get; set; }
        public string Value { get { return value; } set { SetValue(value); } }
        public string Name { get; private set; }

        internal Cell(string name)
        {
            Name = name;
        }

        private void SetValue(string value)
        {
            double numeric;
            if (!double.TryParse(value, out numeric))
            {
                //look at sharedStrings.xml if this one exists and add reference, if not create at sharedStrings.xml and referenced it.
                //XElementCell.SetElementValue(XName.Get("v", XElementCell.GetDefaultNamespace().ToString()), sharedStringIndex);
            }
            else
            {
                //XElementCell.SetElementValue(XName.Get("v", XElementCell.GetDefaultNamespace().ToString()), value);
                XElementCell.SetElementValue(XElementCell.GetDefaultNamespace() + "v", value);
            }
            this.value = value;
        }
    }
}
