using System.Xml.Linq;

namespace ExcelUtility.Impl
{
    public class Drawing
    {
        private XElement drawingData;
        private XNamespace NamespaceWithPrefixXdr { get { return drawingData.GetNamespaceOfPrefix("xdr"); } }
        private XNamespace NamespaceWithPrefixA { get { return drawingData.GetNamespaceOfPrefix("a"); } }
        private IWorksheet parentWorksheet;

        public Drawing(XElement drawingData, IWorksheet worksheet)
        {
            this.drawingData = drawingData;
            parentWorksheet = worksheet;
        }

        public Shape DrawShape(Column column, Row row, double rowOffSet, double colOffSet, double width, double height)
        {
            var from = new XElement(NamespaceWithPrefixXdr + "from", 
                                    new XElement(NamespaceWithPrefixXdr + "col", column.ColumnIndex),
                                    new XElement(NamespaceWithPrefixXdr + "colOff", colOffSet),
                                    new XElement(NamespaceWithPrefixXdr + "row", row),
                                    new XElement(NamespaceWithPrefixXdr + "rowOff", rowOffSet));

            Column destColumn = parentWorksheet.CalculateColumnAfter(column, colOffSet, width);
            Row destRow = parentWorksheet.CalculateRowAfter(row, rowOffSet, height);

            var to = new XElement(NamespaceWithPrefixXdr + "to",
                                    new XElement(NamespaceWithPrefixXdr + "col", destColumn.Name),
                                    new XElement(NamespaceWithPrefixXdr + "colOff", column.Position + colOffSet + width - destColumn.Position), 
                                    new XElement(NamespaceWithPrefixXdr + "row", destRow),
                                    new XElement(NamespaceWithPrefixXdr + "rowOff", row.Position + rowOffSet + height - destRow.Position));
            

            var shape = new XElement(NamespaceWithPrefixXdr + "twoCellAnchor", from, to, 
                new XElement(NamespaceWithPrefixXdr + "sp"), 
                new XElement(NamespaceWithPrefixXdr + "clientData"));
            return null;
        }

        private int GetColumnIndex(string columnReference)
        {
            int index = 0;
            foreach (char c in columnReference)
                index = index + c - 65;
            return index + (26 * (columnReference.Length - 1));
        }
    }
}
