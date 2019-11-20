using System;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using ExcelUtility.UnitTests.Util;
using Xunit;

namespace ExcelUtility.UnitTests.Tests
{
    public class CellTests
    {
        private const string path = @"D:\Projects\Development\CS\SLB\GT\gt-core\Lib\ATTJ_Equipment_List_Template - Copy.xlsx";
        private const string sheetName = "Equipment";
        private ReflectionUtil reflection;

        public CellTests()
        {
            reflection = new ReflectionUtil();
        }

        const double doubleValue = 3.78942147;
        const string stringValue = "Blá!";
        const string fakeStringValue = "123";
        const char charNumber = 'B';

        [Fact]
        public void TestWriteNullValueCell()
        {
            var excelFile = ExcelFile.Open(path);
            IWorksheet worksheet = excelFile.OpenWorksheet(sheetName);
            var cellName = "D19";

            worksheet.GetCell(cellName).StringValue = null;

            excelFile.Save();
            
            Assert.Null(worksheet.GetCell(cellName).StringValue);
        }


        [Fact]
        public void TestWriteStringAtCell()
        {
            var cellName = "D18";
            var cellValue = "Peteca";

            var excelFile = ExcelFile.Open(path);
            IWorksheet worksheet = excelFile.OpenWorksheet(sheetName);
            worksheet.GetCell(cellName).StringValue = cellValue;

            //excelFile.Save();
            //excelFile.Close();

            //excelFile = ExcelFile.Open(path);
            //worksheet = excelFile.OpenWorksheet(sheetName);

            Assert.Equal(cellValue, worksheet.GetCell(cellName).StringValue);
            excelFile.Close();
        }


        [Fact]
        public void TestDoubleValue()
        {

            //put cell names in vars...

            var excelFile = ExcelFile.Open(path);
            IWorksheet worksheet = excelFile.OpenWorksheet(sheetName);

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
            excelFile.Save();
            excelFile.Close();
        }

        [Fact]
        public void TestDoubleValue2()
        {
            var excelFile = ExcelFile.Open(path);
            var worksheet = excelFile.OpenWorksheet(sheetName);

            var sheetData = (XElement)reflection.GetValue(reflection.GetValue(reflection.GetValue(worksheet, "sheetData"), "data"), "data");

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
            excelFile.Save();
            excelFile.Close();

        }

        public void TestStringValue()
        {

        }
    }
}
