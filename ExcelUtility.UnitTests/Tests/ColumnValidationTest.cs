using System.Xml.Linq;
using ExcelUtility.UnitTests.Utility;
using ExcelUtility.Utils;
using Xunit;

namespace ExcelUtility.UnitTests.Tests
{
    public class ColumnValidationTest
    {
        private const double NewWidth = 12;
        private const string path = @"d:\temp\DefaultWorksheet.xlsx";
        private const string sheetName = "Paosdpoasdp";
        private ReflectionUtil reflection;

        public ColumnValidationTest()
        {
            reflection = new ReflectionUtil();
        }

        [Fact]
        public void GetColumn()
        {
            using (var excelFile = ExcelFile.Open(path))
            {
                var worksheet = excelFile.OpenWorksheet(sheetName);
                var colA = worksheet.GetColumn("A");
                Assert.NotNull(colA);
                var cola = worksheet.GetColumn("a");
                Assert.NotNull(cola);
                Assert.Equal(colA, cola);

                var col = worksheet.GetColumn("ZZ");
                Assert.NotNull(col);
                col = worksheet.GetColumn("GH");
                Assert.NotNull(col);

                col = worksheet.GetColumn("AAA");
                Assert.Null(col);
            }
        }

        [Fact]
        public void TestingRangeColumns()
        {
            using (var excelFile = ExcelFile.Open(path))
            {
                var worksheet = excelFile.OpenWorksheet(sheetName);
                worksheet.GetColumn("F").Width = NewWidth;
                worksheet.GetColumn("G").Width = NewWidth;
                worksheet.GetColumn("H").Width = NewWidth;
                worksheet.GetColumn("I").Width = NewWidth;
                worksheet.GetColumn("J").Width = NewWidth;
            }

            using (var excelFile = ExcelFile.Open(path))
            {
                var worksheet = excelFile.OpenWorksheet(sheetName);

                var data = (XElement) reflection.GetValue((XElementData)reflection.GetValue(worksheet, "data"), "data");

                Assert.Equal(NewWidth, worksheet.GetColumn("F").Width);
                Assert.Equal(NewWidth, worksheet.GetColumn("G").Width);
                Assert.Equal(NewWidth, worksheet.GetColumn("H").Width);
                Assert.Equal(NewWidth, worksheet.GetColumn("I").Width);
                Assert.Equal(NewWidth, worksheet.GetColumn("J").Width);
            }

            //Changing midle column of range
            using (var excelFile = ExcelFile.Open(path))
            {
                var worksheet = excelFile.OpenWorksheet(sheetName);
                var column = worksheet.GetColumn("H");
                column.Width = NewWidth / 2;
            }

            using (var excelFile = ExcelFile.Open(path))
            {
                var worksheet = excelFile.OpenWorksheet(sheetName);
                Assert.Equal(NewWidth / 2, worksheet.GetColumn("H").Width);
            }

            //Changing first column of range
            using (var excelFile = ExcelFile.Open(path))
            {
                var worksheet = excelFile.OpenWorksheet(sheetName);
                worksheet.GetColumn("F").Width = NewWidth / 2;
            }

            using (var excelFile = ExcelFile.Open(path))
            {
                var worksheet = excelFile.OpenWorksheet(sheetName);
                Assert.Equal(NewWidth / 2, worksheet.GetColumn("F").Width);
            }

            //Changing last column of range.
            using (var excelFile = ExcelFile.Open(path))
            {
                var worksheet = excelFile.OpenWorksheet(sheetName);
                worksheet.GetColumn("J").Width = NewWidth / 2;
            }

            using (var excelFile = ExcelFile.Open(path))
            {
                var worksheet = excelFile.OpenWorksheet(sheetName);
                Assert.Equal(NewWidth / 2, worksheet.GetColumn("J").Width);
            }
        }

        [Fact]
        public void ChangeCustonWidth()
        {
            //This property, by self 
            using (var excelFile = ExcelFile.Open(path))
            {
                var worksheet = excelFile.OpenWorksheet(sheetName);
                worksheet.GetColumn("I").CustomWidth = 2234;
            }

            using (var excelFile = ExcelFile.Open(path))
            {
                var worksheet = excelFile.OpenWorksheet(sheetName);
                Assert.Equal(2234, worksheet.GetColumn("I").CustomWidth);
            }

            //TODO: Check file integrity.
        }
    }
}
