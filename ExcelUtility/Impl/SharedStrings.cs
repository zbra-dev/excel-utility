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
                unusedStringReferences.Add(new StringReference() { Index = index, OldIndex = index, Text = element.Descendants(Namespace + "t").First().Value });
                stringRefences.Add(new StringReference() { Index = index, OldIndex = index, Text = element.Descendants(Namespace + "t").First().Value });
                index++;
            }
                
            foreach (IWorksheet worksheet in worksheets)
            {
                //send this list to worksheet, the itens founded there are removed from it.
                worksheet.RemoveUnusedStringReferences(unusedStringReferences);
            }
            
            //the remaining items are the items with no reference.
            if (unusedStringReferences.Count > 0)
            {
                //update the referenceString.
                foreach (StringReference unusedStringReference in unusedStringReferences)
                {
                    stringRefences.Remove(unusedStringReference);
                    for (int k = unusedStringReference.Index; k < stringRefences.Count; k++)
                        stringRefences[k].Index--;
                }

                foreach (IWorksheet worksheet in worksheets)
                {
                    //Update references.
                    worksheet.UpdateStringReferences(stringRefences.Where(sr => sr.Index != sr.OldIndex).ToList());
                }
            }

            //Update file.

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
