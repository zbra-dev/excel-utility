using System;
using System.Collections.Generic;
using System.Linq;
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

        public void CleanUpReferences(IList<IWorksheet> worksheets)
        {
            //prepare a list with all references
            IList<StringReference> unusedStringReferences = new List<StringReference>();
            IList<StringReference> stringRefences = new List<StringReference>();

            int index = 0;
            foreach (XElement element in sharedStringsData.Descendants(Namespace + "si"))
            {
                unusedStringReferences.Add(new StringReference() { Reference = element, Index = index, OldIndex = index, Text = element.Descendants(Namespace + "t").First().Value });
                stringRefences.Add(new StringReference() { Reference = element, Index = index, OldIndex = index, Text = element.Descendants(Namespace + "t").First().Value });
                index++;
            }

            //send this list to worksheet, the itens founded there are removed from it.    
            int stringsUsed = 0;
            foreach (IWorksheet worksheet in worksheets)
            {
                worksheet.RemoveUnusedStringReferences(unusedStringReferences);
                stringsUsed += worksheet.CountStringsUsed();
            }

            sharedStringsData.SetAttributeValue("count", stringsUsed);

            //the remaining items are the items with no reference.
            if (unusedStringReferences.Count > 0)
            {
                //update the referenceString.
                foreach (StringReference unusedStringReference in unusedStringReferences)
                {
                    stringRefences.Remove(stringRefences.First(sr => sr.Reference == unusedStringReference.Reference));
                    for (int k = unusedStringReference.Index; k < stringRefences.Count; k++)
                        stringRefences[k].Index--;
                    unusedStringReference.Reference.Remove();
                }
                
                //Update references.
                foreach (IWorksheet worksheet in worksheets)
                    worksheet.UpdateStringReferences(stringRefences.Where(sr => sr.Index != sr.OldIndex).ToList());

                sharedStringsData.SetAttributeValue("uniqueCount", stringRefences.Count);
                //sharedStringsData.SetAttributeValue("uniqueCount", Convert.ToInt32(sharedStringsData.Attribute("uniqueCount").Value) - unusedStringReferences.Count);
            }
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
            XElement newString = new XElement(Namespace + "si", new XElement(Namespace + "t"));
            newString.SetElementValue(Namespace + "t", value);
            
            sharedStringsData.Add(newString);
            //Update count value
            sharedStringsData.SetAttributeValue("count", Convert.ToInt32(sharedStringsData.Attribute("count").Value) + 1);
            sharedStringsData.SetAttributeValue("uniqueCount", Convert.ToInt32(sharedStringsData.Attribute("uniqueCount").Value) + 1);
            return Convert.ToInt32(sharedStringsData.Attribute("uniqueCount").Value) - 1;//Index starts at ZERO!
        }
    }
}
