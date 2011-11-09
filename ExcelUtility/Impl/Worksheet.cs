using System;
using System.Linq;
using System.Xml.Linq;
using System.Xml.XPath;
using System.Text.RegularExpressions;
using System.Collections;
using System.Collections.Generic;
using ExcelUtility.Utils;

namespace ExcelUtility.Impl
{
    internal class Worksheet : IWorksheet
    {
        #region IWorksheet Members

        private XElement sheet;
        private XNamespace Namespace { get { return sheet.GetDefaultNamespace(); } }

        public SharedStrings SharedStrings { get; set; }
        public string Name { get; set; }
        public int SheetId { get; set; }

        internal Worksheet(XElement sheet)
        {
            this.sheet = sheet;
        }

        public Column GetColumn(string name)
        {
            /*
            return (from col in Sheet.Descendants(XName.Get("col", Sheet.GetDefaultNamespace().ToString()))
                    where col.At
                    select new Column(this, name)
                       ).FirstOrDefault();
             */
            return null;
        }

        public Cell GetCell(string name)
        {
            Cell cell = (from c in sheet.Descendants(Namespace + "c")
                         where c.Attribute("r").Value == name
                         select new Cell(c, this, name)).FirstOrDefault();
            return cell == null ? CreateNewCell(name) : cell;
        }

        public Row GetRow(string name)
        {
            throw new NotImplementedException();
        }

        public Shape DrawShape(double x, double y, double width, double height)
        {
            throw new NotImplementedException();
        }

        public void SaveChanges(string xmlPath)
        {
            SharedStrings.Save(xmlPath);
            sheet.Save(string.Format("{0}/worksheets/sheet{1}.xml", xmlPath, SheetId));
        }

        public void RemoveUnusedStringReferences(IList<StringReference> unusedStringRefences)
        {
            IEnumerable<XElement> cells = (from cell in sheet.Descendants(Namespace + "c")
                                           where cell.Attribute("t") != null && cell.Attribute("t").Value == "s"
                                           select cell).ToList();
            foreach (XElement cell in cells)
                if (unusedStringRefences.Any(c => c.Index == Convert.ToInt32(cell.Descendants(Namespace + "v").First().Value)))
                    unusedStringRefences.Remove(unusedStringRefences.First(c => c.Index == Convert.ToInt32(cell.Descendants(Namespace + "v").First().Value)));
        }

        public void UpdateStringReferences(IList<StringReference> stringRefences)
        {
            foreach (StringReference stringReference in stringRefences)
            {
                IList<XElement> cells = (from cell in sheet.Descendants(Namespace + "c")
                                         where cell.Attribute("t") != null && cell.Attribute("t").Value == "s" && Convert.ToInt32(cell.Descendants(Namespace + "v").First().Value) == stringReference.OldIndex
                                         select cell).ToList();
                foreach (XElement cell in cells)
                    cell.Descendants(Namespace + "v").First().Value = stringReference.Index.ToString();
            }
        }

        public int CountStringsUsed()
        {
            return (from cell in sheet.Descendants(Namespace + "c")
                    where cell.Attribute("t") != null && cell.Attribute("t").Value == "s"
                    select cell).ToList().Count;
        }

        #endregion

        private Cell CreateNewCell(string name)
        {
            var newCell = new XElement(Namespace + "c", new XElement(Namespace + "v"));
            newCell.SetAttributeValue("r", name);
            var row = sheet.Descendants(Namespace + "row").Where(r => r.Attribute("r").Value == Regex.Match(name, @"\d+").Value).FirstOrDefault();

            var cells = row.Descendants(Namespace + "c").ToArray();
           
            var cellComparison = new Comparison<XElement>((c1, c2) => 
                {
                    var v1 = c1.Attribute("r").Value;
                    var v2 = c2.Attribute("r").Value;
                    int compare = v1.Length.CompareTo(v2.Length);
                    if (compare == 0)
                        return v1.CompareTo(v2);
                    return compare;
                });

            int index = cells.BinarySearch(newCell, cellComparison);
            if (index < 0)
            {
                index = ~index;
            }
            else
            {
                throw new InvalidOperationException(string.Format("Cell {0} already exists", name));
            }
            if (index >= cells.Length)
                cells[cells.Length - 1].AddAfterSelf(newCell);
            else
                cells[index].AddBeforeSelf(newCell);
            return new Cell(newCell, this, name);
        }
    }
}
