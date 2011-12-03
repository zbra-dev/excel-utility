using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelUtility;
using Xunit;

namespace ExcelUtility.UnitTests.Tests
{
    public class ColumnValidationTest
    {
        private const double NewWidth = 12;
        private const string path = @"d:\temp\DefaultWorksheet.xlsx";
        private const string sheetName = "Paosdpoasdp";


        [Fact]
        public void GetUnexistingColumn()
        {
            using (var excelFile = ExcelFile.Open(path))
            {
                var worksheet = excelFile.OpenWorksheet(sheetName);
                IColumn col = worksheet.GetColumn("GX");
                Assert.NotNull(col);
            }
        }

        [Fact]
        public void GetColumnWithLowerCase()
        {
            using (var excelFile = ExcelFile.Open(path))
            {
                var worksheet = excelFile.OpenWorksheet(sheetName);
                IColumn col = worksheet.GetColumn("m");
                Assert.NotNull(col);
            }
        }

        [Fact]
        public void CheckingWidth()
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
                Assert.Equal(NewWidth, worksheet.GetColumn("F").Width);
                Assert.Equal(NewWidth, worksheet.GetColumn("G").Width);
                Assert.Equal(NewWidth, worksheet.GetColumn("H").Width);
                Assert.Equal(NewWidth, worksheet.GetColumn("I").Width);
                Assert.Equal(NewWidth, worksheet.GetColumn("J").Width);
            }
        }

        [Fact]
        public void ChangeMidleColumnInRange()
        {
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
        }

        [Fact]
        public void ChangeFirstColumnInRange()
        {
            using (var excelFile = ExcelFile.Open(path))
            {
                var worksheet = excelFile.OpenWorksheet(sheetName);
                var firstColumn = worksheet.GetColumn("F");
                firstColumn.Width = NewWidth / 2;
            }

            using (var excelFile = ExcelFile.Open(path))
            {
                var worksheet = excelFile.OpenWorksheet(sheetName);
                Assert.Equal(NewWidth / 2, worksheet.GetColumn("F").Width);
            }
        }

        [Fact]
        public void ChangeLastColumnInRange()
        {
            using (var excelFile = ExcelFile.Open(path))
            {
                var worksheet = excelFile.OpenWorksheet(sheetName);
                var lastColumn = worksheet.GetColumn("J");
                lastColumn.Width = NewWidth / 2;
            }

            using (var excelFile = ExcelFile.Open(path))
            {
                var worksheet = excelFile.OpenWorksheet(sheetName);
                Assert.Equal(NewWidth / 2, worksheet.GetColumn("J").Width);
            }

        }

        [Fact]
        public void FakeChangeAtColumn()
        {
            using (var excelFile = ExcelFile.Open(path))
            {
                var worksheet = excelFile.OpenWorksheet(sheetName);
                var anyColumn = worksheet.GetColumn("I");
                anyColumn.Width = NewWidth;
            }

            using (var excelFile = ExcelFile.Open(path))
            {
                var worksheet = excelFile.OpenWorksheet(sheetName);
                Assert.Equal(NewWidth, worksheet.GetColumn("I").Width);
            }
        }
    }
}
