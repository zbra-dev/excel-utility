using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Xunit;
using ExcelUtility.UnitTests.Util;
using ExcelUtility.Utils;

namespace ExcelUtility.UnitTests.Tests
{
    public class XElementDataTest
    {
        private const string path = "";
        private const string sheetName = "";
        private ReflectionUtil reflection;

        public XElementDataTest()
        {
            reflection = new ReflectionUtil();
        }

        [Fact]
        public void getDescendants()
        {
            using (var file = ExcelFile.Open(path))
            {
                var sheet = file.OpenWorksheet(sheetName);
                var sheetXElementData = (XElementData)reflection.GetValue(sheet, "data");
                Assert.NotNull(sheetXElementData);

                //Get main descedants
                var sheetViews = sheetXElementData.Descendants("sheetViews");
                Assert.NotNull(sheetViews);
                Assert.NotNull(sheetViews.FirstOrDefault());
                var sheetFormatPr = sheetXElementData.Descendants("sheetFormatPr");
                Assert.NotNull(sheetFormatPr);
                Assert.NotNull(sheetFormatPr.FirstOrDefault());
                var cols = sheetXElementData.Descendants("cols");
                Assert.NotNull(cols);
                Assert.NotNull(cols.FirstOrDefault());
                var sheetData = sheetXElementData.Descendants("sheetData");
                Assert.NotNull(sheetData);
                Assert.NotNull(sheetData.FirstOrDefault());
                var pageMargins = sheetXElementData.Descendants("pageMargins");
                Assert.NotNull(pageMargins);
                Assert.NotNull(pageMargins.FirstOrDefault());
                var drawing = sheetXElementData.Descendants("drawing");
                Assert.NotNull(drawing);
                Assert.NotNull(drawing.FirstOrDefault()); 

                //Get child descendents directly
                var sheetView = sheetXElementData.Descendants("sheetView");
                Assert.NotNull(sheetView);
                Assert.NotNull(sheetView.FirstOrDefault());
                var col = sheetXElementData.Descendants("col");
                Assert.NotNull(col);
                Assert.NotNull(col.FirstOrDefault());
                var row = sheetXElementData.Descendants("row");
                Assert.NotNull(row);
                Assert.NotNull(row.FirstOrDefault());

                //Get sub child descendants directly
                var cell = sheetXElementData.Descendants("c");
                Assert.NotNull(cell);
                Assert.NotNull(cell.FirstOrDefault());

                var value = sheetXElementData.Descendants("v");
                Assert.NotNull(value);
                Assert.NotNull(value.FirstOrDefault());

                //Get unExisting descedants
                Assert.Throws<Exception>(() => sheetXElementData.Descendants("nonExisting"));
            }
        }

        public void addingContent()
        {
            using (var file = ExcelFile.Open(path))
            {
                var sheet = file.OpenWorksheet(sheetName);
                var sheetXElementData = (XElementData)reflection.GetValue(sheet, "data");
                Assert.NotNull(sheetXElementData);

                sheetXElementData.Add("newElement");

            }

        }
    }
}
