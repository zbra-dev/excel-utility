using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Xunit;
using System.Globalization;
using ExcelUtility.UnitTests.Util;
using System.Xml.Linq;

namespace ExcelUtility.UnitTests.Tests
{
    public class CellTests
    {
        private string path = @"D:/temp/DefaultWorksheet.xlsx";
        private string sheetName = "Paosdpoasdp";
        private ReflectionUtil reflection;

        public CellTests()
        {
            reflection = new ReflectionUtil();
        }

        [Fact]
        public void TestDoubleValue()
        {
            const double doubleValue = 3.78942147;
            const string stringValue = "Blá!";
            const string fakeStringValue = "123";
            const char charNumber = 'B';

            //put cell names in vars...

            using (var excelFile = ExcelFile.Open(path))
            {
                IWorksheet worksheet = excelFile.OpenWorksheet(sheetName);

                Assert.DoesNotThrow(() => worksheet.GetCell("c3"));
                var cellLowerCase = worksheet.GetCell("c3");
                Assert.NotNull(cellLowerCase);

                var doubleValueCell = worksheet.GetCell("A4");
                Assert.NotNull(doubleValueCell);
                doubleValueCell.DoubleValue = doubleValue;

                var stringValueCell = worksheet.GetCell("A5");
                Assert.NotNull(stringValueCell);
                stringValueCell.StringValue = stringValue;

                var fakeStringValueCell = worksheet.GetCell("A6");
                Assert.NotNull(fakeStringValueCell);
                fakeStringValueCell.StringValue = fakeStringValue;

                var charNumberValueCell = worksheet.GetCell("A7");
                Assert.NotNull(charNumberValueCell);
                charNumberValueCell.DoubleValue = charNumber;
            }

            using (var excelFile = ExcelFile.Open(path))
            {
                IWorksheet worksheet = excelFile.OpenWorksheet(sheetName);

                var sheetData = (XElement)reflection.GetValue(reflection.GetValue(worksheet, "sheetData"), "data");
                var cellList = sheetData.Descendants(sheetData.GetDefaultNamespace() + "c");

                var doubleValueCell = cellList.Where(c => c.Attribute("r").Value == "A4").FirstOrDefault();
                Assert.NotNull(doubleValueCell);
                Assert.Equal(doubleValue, Convert.ToDouble(doubleValueCell.Descendants(sheetData.GetDefaultNamespace() + "v").First().Value, NumberFormatInfo.InvariantInfo));

                //Certificate that it is a string inside.
                var stringValueCell = cellList.Where(c => c.Attribute("r").Value == "A5").FirstOrDefault();
                Assert.NotNull(stringValueCell);
                var tProperty = stringValueCell.Attribute("t");
                Assert.NotNull(tProperty);
                Assert.Equal("s", tProperty.Value);
                //Check if value isn't the inserted string.
                Assert.NotEqual(stringValue, stringValueCell.Descendants(sheetData.GetDefaultNamespace() + "v").First().Value);

                //Certificate that it is a string inside.
                var fakeNumberValueCell = cellList.Where(c => c.Attribute("r").Value == "A6").FirstOrDefault();
                Assert.NotNull(stringValueCell);
                var fakeTProperty = stringValueCell.Attribute("t");
                Assert.NotNull(fakeTProperty);
                Assert.Equal("s", fakeTProperty.Value);
                //Check if value isn't the inserted fake string.
                Assert.NotEqual(fakeStringValue, stringValueCell.Descendants(sheetData.GetDefaultNamespace() + "v").First().Value);

                var charNumberValueCell = cellList.Where(c => c.Attribute("r").Value == "A7").FirstOrDefault();
                Assert.NotNull(charNumberValueCell);
                Assert.Equal(charNumber, Convert.ToInt32(charNumberValueCell.Descendants(sheetData.GetDefaultNamespace() + "v").First().Value));
            }
        }

        public void TestStringValue()
        {

        }
    }
}
