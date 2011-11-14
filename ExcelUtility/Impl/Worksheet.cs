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
        private XElement sheetData;
        private XElement relationshipsData;
        private string worksheetPath;
        private XNamespace SheetNamespace { get { return sheetData.GetDefaultNamespace(); } }
        private XNamespace SheetNamespaceWithPrefixR { get { return sheetData.GetNamespaceOfPrefix("r"); } }
        private XNamespace RelationshipNamespace { get { return relationshipsData.GetDefaultNamespace(); } }
        
        internal Worksheet(XElement sheet, XElement sheetRelationships, string path)
        {
            sheetData = sheet;
            relationshipsData = sheetRelationships;
            worksheetPath = path;
            LoadDrawing();
        }

        #region IWorksheet Members
        public Drawing Drawing { get; set; }
        public SharedStrings SharedStrings { get; set; }
        public string Name { get; set; }
        public int SheetId { get; set; }
        public double DefaultRowHeight { get { return Convert.ToDouble(sheetData.Descendants(SheetNamespace + "sheetFormatPr").First().Attribute("defaultRowHeight").Value); } }

        public Column GetColumn(string name)
        {
            var col = (from c in sheetData.Descendants(SheetNamespace + "col")
                       where Convert.ToInt32(c.Attribute("min").Value) >= GetColumnIndex(name) && Convert.ToInt32(c.Attribute("max").Value) <= GetColumnIndex(name)
                       select new Column(c, this, name)
                       {
                           Width = Convert.ToDouble(c.Attribute("width").Value)
                       }).FirstOrDefault();
            return col == null ? CreateNewColumn(name) : col;
        }

        public Column CalculateColumnAfter(Column columnBase, double colOffSet, double width)
        {
            double dif = colOffSet;
            int colCount = 0;
            var cols = sheetData.Descendants(SheetNamespace + "col").Where(r => Convert.ToInt32(r.Attribute("r").Value) > columnBase.ColumnIndex).ToList();
            while (dif > 0)
            {
                if (colCount > cols.Count)
                    dif -= DefaultRowHeight;
                else
                    dif -= Convert.ToDouble(cols[colCount].Attribute("width").Value);
                colCount++;
            }
            return colCount > cols.Count ? CreateNewColumn(GetColumnNameBy(colCount)) : new Column(cols[colCount], this, GetColumnNameBy(colCount)) { Width = Convert.ToDouble(cols[colCount].Attribute("widht").Value) };
        }

        public double GetColumnPosition(int columnIndex)
        {
            double colPosition = 0;
            /*var cols = (from col in sheetData.Descendants(SheetNamespace + "col")
                        where (Convert.ToInt32(col.Attribute("max").Value) <= columnIndex || (Convert.ToInt32(col.Attribute("min").Value) < columnIndex && Convert.ToInt32(col.Attribute("max").Value) > columnIndex))
                        select new Column(col, this, "") { Width = Convert.ToDouble(col.Attribute("width").Value) }).ToList();
             */
            var cols = new List<XElement>();

            //All data
            var colsData = (from col in sheetData.Descendants(SheetNamespace + "col")
                            where (Convert.ToInt32(col.Attribute("max").Value) <= columnIndex || (Convert.ToInt32(col.Attribute("min").Value) < columnIndex && Convert.ToInt32(col.Attribute("max").Value) > columnIndex))
                            select col).ToList();

            //Just Xelements that represent a range
            var rangeCols = colsData.Where(t => t.Attribute("min").Value != t.Attribute("max").Value);
            foreach (var col in rangeCols)
            {
                int range = Convert.ToInt32(col.Attribute("max").Value) - Convert.ToInt32(col.Attribute("min").Value);
                for (int k = 0; k < range; k++)
                {

                }
            }

            //Just XElement that doesn't represent a range
            cols.AddRange(colsData.Where(t => t.Attribute("min").Value == t.Attribute("max").Value));

            //foreach (var col in cols)
            //    if(col.ColumnIndex <= columnIndex)
            //        colPosition += col.Width;
            return colPosition;
        }

        public Row GetRow(int index)
        {
            var row = (from r in sheetData.Descendants(SheetNamespace + "row")
                       where Convert.ToInt32(r.Attribute("r").Value) == index
                       select new Row(r, this, index)
                       {
                           Height = r.Attribute("ht") == null ? DefaultRowHeight : Convert.ToDouble(r.Attribute("ht").Value)
                       }).FirstOrDefault();
            return row == null ? CreateNewRow(index) : row;
        }

        public Row CalculateRowAfter(Row row, double rowOffSet, double height)
        {
            double dif = rowOffSet;
            int rowCount = 0;
            var rows = sheetData.Descendants(SheetNamespace + "row").Where(r => Convert.ToInt32(r.Attribute("r").Value) > row.Index).ToList();
            while (dif > 0){
                if (rowCount > rows.Count)
                    dif -= DefaultRowHeight;
                else
                    dif -= rows[rowCount].Attribute("ht") == null ? DefaultRowHeight : Convert.ToDouble(rows[rowCount].Attribute("ht").Value);
                rowCount++;
            }
            return rowCount > rows.Count ? CreateNewRow(rowCount) : new Row(rows[rowCount], this, rowCount) { Height = rows[rowCount].Attribute("ht") == null ? DefaultRowHeight : Convert.ToDouble(rows[rowCount].Attribute("ht").Value) };
        }

        public double GetRowPosition(int rowIndex)
        {
            double rowPosition = 0;
            var rowsHeight = (from row in sheetData.Descendants(SheetNamespace + "row")
                                        where Convert.ToInt32(row.Attribute("r").Value) < rowIndex
                                        select row.Attribute("ht") == null ? DefaultRowHeight : Convert.ToDouble(row.Attribute("ht").Value)).ToList();
            foreach (var rowHeight in rowsHeight)
                rowPosition += rowHeight;
            return rowPosition;
        }

        public Cell GetCell(string name)
        {
            Cell cell = (from c in sheetData.Descendants(SheetNamespace + "c")
                         where c.Attribute("r").Value == name
                         select new Cell(c, this, name)).FirstOrDefault();
            return cell == null ? CreateNewCell(name) : cell;
        }

        public Shape DrawShape(string columnReference, int rowReference, double xOffSet, double yOffSet, double width, double height)
        {
            //return Drawing.DrawShape(columnReference, rowReference, xOffSet, yOffSet, width, height);            
            return null;
        }

        public void SaveChanges(string xmlPath)
        {
            SharedStrings.Save(xmlPath);
            sheetData.Save(string.Format("{0}/worksheets/sheet{1}.xml", xmlPath, SheetId));
        }

        public void RemoveUnusedStringReferences(IList<StringReference> unusedStringRefences)
        {
            IEnumerable<XElement> cells = (from cell in sheetData.Descendants(SheetNamespace + "c")
                                           where cell.Attribute("t") != null && cell.Attribute("t").Value == "s"
                                           select cell).ToList();
            foreach (XElement cell in cells)
                if (unusedStringRefences.Any(c => c.Index == Convert.ToInt32(cell.Descendants(SheetNamespace + "v").First().Value)))
                    unusedStringRefences.Remove(unusedStringRefences.First(c => c.Index == Convert.ToInt32(cell.Descendants(SheetNamespace + "v").First().Value)));
        }

        public void UpdateStringReferences(IList<StringReference> stringRefences)
        {
            foreach (StringReference stringReference in stringRefences)
            {
                IList<XElement> cells = (from cell in sheetData.Descendants(SheetNamespace + "c")
                                         where cell.Attribute("t") != null && cell.Attribute("t").Value == "s" && Convert.ToInt32(cell.Descendants(SheetNamespace + "v").First().Value) == stringReference.OldIndex
                                         select cell).ToList();
                foreach (XElement cell in cells)
                    cell.Descendants(SheetNamespace + "v").First().Value = stringReference.Index.ToString();
            }
        }

        public int CountStringsUsed()
        {
            return (from cell in sheetData.Descendants(SheetNamespace + "c")
                    where cell.Attribute("t") != null && cell.Attribute("t").Value == "s"
                    select cell).ToList().Count;
        }

        #endregion
        
        private void LoadDrawing()
        {
            var drawing = sheetData.Descendants(SheetNamespace + "drawing").FirstOrDefault();
            if (drawing != null)
            {
                string drawingId = drawing.Attribute(SheetNamespaceWithPrefixR + "id").Value;
                //Load 
                var drawRelationship = relationshipsData.Descendants(RelationshipNamespace + "Relationship").Where(r => r.Attribute("Id").Value == drawingId).First();
                Drawing = new Drawing(XElement.Load(string.Format("{0}/{1}", worksheetPath, drawRelationship.Attribute("Target").Value)), this);
            }
        }
        
        private Cell CreateNewCell(string name)
        {
            var newCell = new XElement(SheetNamespace + "c", new XElement(SheetNamespace + "v"));
            newCell.SetAttributeValue("r", name);

            var row = sheetData.Descendants(SheetNamespace + "row").Where(r => r.Attribute("r").Value == Regex.Match(name, @"\d+").Value).FirstOrDefault();
            var cells = row.Descendants(SheetNamespace + "c").ToArray();
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
            index = ~index;
            if (index >= cells.Length)
                cells[cells.Length - 1].AddAfterSelf(newCell);
            else
                cells[index].AddBeforeSelf(newCell);
            return new Cell(newCell, this, name);
        }

        private Row CreateNewRow(int rowCount)
        {
            var newRow = new XElement(SheetNamespace + "row");
            newRow.SetAttributeValue("r", rowCount);

            var rows = sheetData.Descendants(SheetNamespace + "row").ToArray();
            var rowComparison = new Comparison<XElement>((r1, r2) =>
            {
                var v1 = r1.Attribute("r").Value;
                var v2 = r2.Attribute("r").Value;
                int compare = v1.Length.CompareTo(v2.Length);
                if (compare == 0)
                    return v1.CompareTo(v2);
                return compare;
            });
            int index = rows.BinarySearch(newRow, rowComparison);
            index = ~index;
            if (index >= rows.Length)
                rows[rows.Length - 1].AddAfterSelf(newRow);
            else
                rows[index].AddBeforeSelf(newRow);

            return new Row(newRow, this, rowCount) { Height = DefaultRowHeight };
        }

        private Column CreateNewColumn(string name)
        {
            return null;
        }

        private string GetColumnNameBy(int colCount)
        {
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