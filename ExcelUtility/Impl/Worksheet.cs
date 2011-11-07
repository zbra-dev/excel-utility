using System;
using System.Linq;
using System.Xml.Linq;
using System.Text.RegularExpressions;
using System.Collections;

namespace ExcelUtility.Impl
{
    internal class Worksheet : IWorksheet
    {
        #region IWorksheet Members
        public XElement Sheet { get; set; }
        public string Name { get; set; }
        public int SheetId { get; set; }
        
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
            Cell cell = (from c in Sheet.Descendants(Sheet.GetDefaultNamespace() + "c")
                         where c.Attribute("r").Value == name
                         select new Cell(name)
                         {
                             XElementCell = c
                          }).FirstOrDefault();
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
            Sheet.Save(string.Format("{0}/sheet{1}.xml", xmlPath, SheetId));
        }

        #endregion

        private Cell CreateNewCell(string name)
        {
            XElement newCell = new XElement(Sheet.GetDefaultNamespace() + "c", new XElement(Sheet.GetDefaultNamespace() + "v"));
            newCell.SetAttributeValue(Sheet.GetDefaultNamespace() + "r", name);
            XElement row = Sheet.Descendants(Sheet.GetDefaultNamespace() + "row").Where(r => r.Attribute("r").Value == Regex.Match(name, @"\d+").Value).FirstOrDefault();
            //Check if row is null.
            row.Add(newCell);
            return new Cell(name) { XElementCell = newCell };
        }
    }
}
