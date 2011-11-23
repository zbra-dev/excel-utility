using System;
using System.Linq;
using System.Xml.Linq;
using System.Xml.XPath;
using System.Text.RegularExpressions;
using System.Collections;
using System.Collections.Generic;
using ExcelUtility.Utils;
using System.Globalization;

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
        public double DefaultRowHeight { get { return Convert.ToDouble(sheetData.Descendants(SheetNamespace + "sheetFormatPr").First().Attribute("defaultRowHeight").Value, CultureInfo.InvariantCulture); } }
        public double DefaultColumnWidth { get { return 9; } }

        #region ColumnArea

        public Column GetColumn(string name)
        {
            var col = (from c in sheetData.Descendants(SheetNamespace + "col")
                       where GetColumnIndex(name) + 1 >= Convert.ToInt32(c.Attribute("min").Value) && GetColumnIndex(name) + 1 <= Convert.ToInt32(c.Attribute("max").Value)
                       select new Column(c, this, name)
                       {
                           Width = Convert.ToDouble(c.Attribute("width").Value, CultureInfo.InvariantCulture)
                       }).FirstOrDefault();
            return col == null ? CreateColumnBetween(GetColumnIndex(name) + 1, GetColumnIndex(name) + 1) : col;
        }

        public Column CalculateColumnAfter(Column columnBase, double colOffSet, double width)
        {
            double dif = width;
            int colCount = 0;
            var colsWidhtValue = new List<double>();

            var cols = sheetData.Descendants(SheetNamespace + "col").Where(col => columnBase.ColumnIndex + 1 <= Convert.ToInt32(col.Attribute("max").Value)).ToList();

            var separatedCols = ((from col in cols
                                     where col.Attribute("min").Value == col.Attribute("max").Value
                                     select col).ToList());

            var colsInRange = cols.Where(t => t.Attribute("min").Value != t.Attribute("max").Value);
            foreach (var col in colsInRange)
            {
                int baseIndex = Convert.ToInt32(col.Attribute("min").Value);
                int range = Convert.ToInt32(col.Attribute("max").Value) - ((columnBase.ColumnIndex + 1 > Convert.ToInt32(col.Attribute("min").Value)) ? columnBase.ColumnIndex + 1 : Convert.ToInt32(col.Attribute("min").Value)) + 1;
                for (int k = 0; k < range; k++)
                    separatedCols.Insert(baseIndex + k - 1, new XElement(SheetNamespace + "col", new XAttribute("min", baseIndex + k), new XAttribute("max", baseIndex + k), new XAttribute("width", col.Attribute("width").Value)));
            }

            while (dif > 0)
            {
                if (colCount > separatedCols.Count)
                    dif -= DefaultColumnWidth;
                else
                    dif -= Convert.ToDouble(separatedCols[colCount].Attribute("width").Value, CultureInfo.InvariantCulture);// colsWidhtValue[colCount];
                colCount++;
            }

            return colCount > separatedCols.Count ? CreateColumnBetween(colCount, colCount) : new Column(separatedCols[colCount - 1], this, GetColumnNameBy(Convert.ToInt32(separatedCols[colCount - 1].Attribute("min").Value))) { Width = Convert.ToDouble(separatedCols[colCount - 1].Attribute("width").Value, CultureInfo.InvariantCulture) };
        }

        public Column CreateColumnBetween(int min, int max)
        {
            return new Column(CreateColumnData(min, max, DefaultColumnWidth), this, min == max ? GetColumnNameBy(min) : "ColumnRange");
        }

        public Column CreateColumnBetweenWith(int min, int max, double width)
        {
            return new Column(CreateColumnData(min, max, width), this, min == max ? GetColumnNameBy(min) : "ColumnRange");
        }

        public double GetColumnPosition(int columnIndex)
        {
            double colPosition = 0;
            var cols = new List<XElement>();
            var colsWidhtValue = new List<Double>();
            //All data
            var colsData = (from col in sheetData.Descendants(SheetNamespace + "col")
                            where (Convert.ToInt32(col.Attribute("max").Value) < columnIndex || (Convert.ToInt32(col.Attribute("min").Value) < columnIndex && Convert.ToInt32(col.Attribute("max").Value) >= columnIndex))
                            select col).ToList();

            //Just XElement that doesn't represent a range
            colsWidhtValue.AddRange((from col in colsData
                                     where col.Attribute("min").Value == col.Attribute("max").Value
                                     select Convert.ToDouble(col.Attribute("width").Value, CultureInfo.InvariantCulture)).ToList());

            //Just Xelements that represent a range
            var rangeCols = colsData.Where(t => t.Attribute("min").Value != t.Attribute("max").Value);
            foreach (var col in rangeCols)
            {
                int baseIndex = Convert.ToInt32(col.Attribute("min").Value) ;
                int range = (Convert.ToInt32(col.Attribute("max").Value) >= columnIndex ? columnIndex : Convert.ToInt32(col.Attribute("max").Value)) - Convert.ToInt32(col.Attribute("min").Value) + 1;
                for (int k = 0; k < range; k++)
                    if (baseIndex + k < columnIndex)
                        colsWidhtValue.Add(Convert.ToDouble(col.Attribute("width").Value, CultureInfo.InvariantCulture));
            }

            //Calculate Width for non listed columns
            colPosition += DefaultColumnWidth * (columnIndex - 1 - colsWidhtValue.Count);

            foreach (var width in colsWidhtValue)
                colPosition += width;
            return colPosition;
        }
        #endregion

        public Row GetRow(int index)
        {
            var row = (from r in sheetData.Descendants(SheetNamespace + "row")
                       where Convert.ToInt32(r.Attribute("r").Value) == index
                       select new Row(r, this, index)
                       {
                           Height = r.Attribute("ht") == null ? DefaultRowHeight : Convert.ToDouble(r.Attribute("ht").Value, CultureInfo.InvariantCulture)
                       }).FirstOrDefault();
            return row == null ? CreateNewRow(index) : row;
        }

        public Row CalculateRowAfter(Row row, double rowOffSet, double height)
        {
            double dif = height;
            int rowCount = 0;
            var rows = sheetData.Descendants(SheetNamespace + "row").Where(r => Convert.ToInt32(r.Attribute("r").Value) >= row.Index).ToList();
            while (dif > 0)
            {
                dif -= rowCount > rows.Count ? DefaultRowHeight : rows[rowCount].Attribute("ht") == null ? DefaultRowHeight : Convert.ToDouble(rows[rowCount].Attribute("ht").Value, CultureInfo.InvariantCulture);
                rowCount++;
            }
            //TODO: check CreateNewRow method.
            return rowCount > rows.Count ? CreateNewRow(rowCount) : new Row(rows[rowCount - 1], this, row.Index + rowCount - 1) { Height = rows[rowCount - 1].Attribute("ht") == null ? DefaultRowHeight : Convert.ToDouble(rows[rowCount - 1].Attribute("ht").Value, CultureInfo.InvariantCulture) };
        }

        public double GetRowPosition(int rowIndex)
        {
            double rowPosition = 0;
            var rowsHeight = (from row in sheetData.Descendants(SheetNamespace + "row")
                              where Convert.ToInt32(row.Attribute("r").Value) < rowIndex
                              select row.Attribute("ht") == null ? DefaultRowHeight : Convert.ToDouble(row.Attribute("ht").Value, CultureInfo.InvariantCulture)).ToList();
            //Sum a DefaultRowHeight measure for each non listed item.
            rowPosition += DefaultRowHeight * (rowIndex - 1 - rowsHeight.Count);
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
            Drawing.Save();
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
                Drawing = new Drawing(string.Format("{0}/{1}", worksheetPath, drawRelationship.Attribute("Target").Value), this);
            }
        }

        private Cell CreateNewCell(string name)
        {
            var newCell = new XElement(SheetNamespace + "c", new XElement(SheetNamespace + "v"));
            newCell.SetAttributeValue("r", name);

            var row = sheetData.Descendants(SheetNamespace + "row").Where(r => r.Attribute("r").Value == Regex.Match(name, @"\d+").Value).FirstOrDefault();
            var cells = row.Descendants(SheetNamespace + "c").ToArray();
            if (cells.Length == 0)
            {
                row.Add(newCell);
            }
            else
            {
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
            }
            return new Cell(newCell, this, name);
        }

        private Row CreateNewRow(int rowCount)
        {
            var newRow = new XElement(SheetNamespace + "row");
            newRow.SetAttributeValue("r", rowCount);

            var rows = sheetData.Descendants(SheetNamespace + "row").ToArray();
            if (rows.Length == 0)
            {
                sheetData.Add(newRow);
            }
            else
            {
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
            }
            return new Row(newRow, this, rowCount) { Height = DefaultRowHeight };
        }

        private string GetColumnNameBy(int colCount)
        {
            if (colCount / 26 == 0)
                return string.Format("{0}", Convert.ToChar(colCount % 26 + 64));
            return string.Format("{0}{1}", Convert.ToChar(colCount / 26 + 64), Convert.ToChar(colCount % (26 * (colCount / 26)) + 64));
        }

        private int GetColumnIndex(string columnReference)
        {
            int index = 0;
            foreach (char c in columnReference)
                index = index + c - 65;
            return index + (26 * (columnReference.Length - 1));
        }

        private XElement CreateColumnData(int min, int max, double width)
        {
            XElement newColumn = new XElement(SheetNamespace + "col");
            newColumn.SetAttributeValue("min", min);
            newColumn.SetAttributeValue("max", max);
            newColumn.SetAttributeValue("width", width);
            newColumn.SetAttributeValue("customWidth", 1);

            var columns = sheetData.Descendants(SheetNamespace + "cols").FirstOrDefault();
            if (columns == null)
            {
                columns = new XElement(SheetNamespace + "cols");
                sheetData.Descendants(SheetNamespace + "sheetFormatPr").First().AddAfterSelf(columns);
            }
            var cols = sheetData.Descendants(SheetNamespace + "col").ToArray();

            if (cols.Length == 0)
            {
                columns.Add(newColumn);
            }
            else
            {
                var colComparison = new Comparison<XElement>((r1, r2) =>
                {
                    var v1 = r1.Attribute("min").Value;
                    var v2 = r2.Attribute("max").Value;
                    return v2.CompareTo(v2);
                });
                int index = cols.BinarySearch(newColumn, colComparison);
                index = ~index;
                if (index >= cols.Length)
                    cols[cols.Length - 1].AddAfterSelf(newColumn);
                else
                    cols[index].AddBeforeSelf(newColumn);
            }            

            return newColumn;
        }
    }
}