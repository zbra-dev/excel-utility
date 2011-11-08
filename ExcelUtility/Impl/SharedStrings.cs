using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace ExcelUtility.Impl
{
    public class SharedStrings
    {
        private XElement sharedStringsData;
        private XNamespace Namespace { get { return sharedStringsData.GetDefaultNamespace(); } }

        public SharedStrings(XElement sharedStringsData)
        {
            this.sharedStringsData = sharedStringsData;
        }

        public void Save(string xmlPath)
        {
            sharedStringsData.Save(string.Format("{0}/sharedStrings.xml",xmlPath));
        }

        public int GetStringReferenceOf(string value)
        {
            if (sharedStringsData.Descendants(Namespace + "t").Any(t => t.Value == value))
            {
                IEnumerable<XElement> elements = sharedStringsData.Descendants(Namespace + "si");
                int index = 0;
                foreach (XElement element in elements)
                {
                    if (element.Descendants(Namespace + "t").FirstOrDefault().Value == value)
                        break;
                    index++;
                }
                return index;
            }
            return CreateNewStringReference(value);
        }

        private int CreateNewStringReference(string value)
        {
            XElement newString = new XElement(Namespace + "si", new XElement(Namespace + "t"), new XElement(Namespace+"s"));
            newString.SetElementValue(Namespace + "t", value);
            newString.SetElementValue(Namespace + "s", 0);
            sharedStringsData.Add(newString);
            //Update count value
            sharedStringsData.SetAttributeValue("uniqueCount", Convert.ToInt32(sharedStringsData.Attribute("uniqueCount").Value) + 1);
            return Convert.ToInt32(sharedStringsData.Attribute("uniqueCount").Value);
        }
    }
}
