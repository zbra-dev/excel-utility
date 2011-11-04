using System;
using System.Xml.Linq;

namespace ExcelUtility.Impl
{
    internal class Worksheet : IWorksheet
    {
        #region IWorksheet Members

        public string Name { get; set; }
        
        public int SheetId { get; set; }
        
        public Column GetColumn(string name)
        {
            throw new NotImplementedException();
        }

        public Cell GetCell(string name)
        {
            throw new NotImplementedException();
        }

        public Row GetRow(string name)
        {
            throw new NotImplementedException();
        }

        public Shape DrawShape(double x, double y, double width, double height)
        {
            throw new NotImplementedException();
        }

        public void SaveChanges(string worksheetsPath)
        {
            XDocument worksheet = XDocument.Load(string.Format("{0}/sheet{1}.xml", worksheetsPath, SheetId));
        }

        #endregion
    }
}
