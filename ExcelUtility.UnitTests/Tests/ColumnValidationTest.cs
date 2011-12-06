using System.Linq;
using System.Xml.Linq;
using ExcelUtility.UnitTests.Util;
using ExcelUtility.Utils;
using Xunit;
using System;
using System.Globalization;

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
                var columnData = (XElement)reflection.GetValue(reflection.GetValue(reflection.GetValue(worksheet, "sheetColumns"), "data"), "data");
                Assert.NotNull(columnData);

                var colA = worksheet.GetColumn("A");
                Assert.NotNull(colA);
                var cola = worksheet.GetColumn("a");
                Assert.NotNull(cola);
                Assert.Equal(colA, cola);

                string lastColumnName = "ZZ";
                var col = worksheet.GetColumn(lastColumnName);
                Assert.NotNull(col);

                var lastColumn = columnData.Descendants(columnData.GetDefaultNamespace() + "col").Where(r => Convert.ToInt64(r.Attribute("min").Value) >= ColumnUtil.GetColumnIndex(lastColumnName) && Convert.ToInt64(r.Attribute("max").Value) <= ColumnUtil.GetColumnIndex(lastColumnName)).FirstOrDefault();
                Assert.NotNull(lastColumn);
                
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
                var data = (XElement)reflection.GetValue(reflection.GetValue(worksheet, "data"), "data");
                
                var rangeCol = data.Descendants(data.GetDefaultNamespace() + "col").Where(c => Convert.ToInt32(c.Attribute("min").Value) == ColumnUtil.GetColumnIndex("F") + 1 && Convert.ToInt32(c.Attribute("max").Value) == ColumnUtil.GetColumnIndex("J") + 1);
                Assert.NotNull(rangeCol);
                
                var item = rangeCol.FirstOrDefault();
                Assert.NotNull(item);

                Assert.Equal(NewWidth, Convert.ToDouble(item.Attribute("width").Value, NumberFormatInfo.InvariantInfo));
                
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
                var data = (XElement)reflection.GetValue(reflection.GetValue(worksheet, "data"), "data");
                var colH = data.Descendants(data.GetDefaultNamespace() + "col").Where(c => Convert.ToInt32(c.Attribute("min").Value) == ColumnUtil.GetColumnIndex("H") && Convert.ToInt32(c.Attribute("max").Value) == ColumnUtil.GetColumnIndex("H")).FirstOrDefault();
                Assert.NotNull(colH);

                Assert.Equal(NewWidth / 2, Convert.ToDouble(colH.Attribute("width").Value, NumberFormatInfo.InvariantInfo));
                
                Assert.Equal(NewWidth / 2, worksheet.GetColumn("H").Width);
            }

            //Changing first column of range
            using (var excelFile = ExcelFile.Open(path))
            {
                var worksheet = excelFile.OpenWorksheet(sheetName);
                worksheet.GetColumn("F").Width = NewWidth / 2 + 2;
            }

            using (var excelFile = ExcelFile.Open(path))
            {
                var worksheet = excelFile.OpenWorksheet(sheetName);

                var data = (XElement)reflection.GetValue(reflection.GetValue(worksheet, "data"), "data");
                var colF = data.Descendants(data.GetDefaultNamespace() + "col").Where(c => Convert.ToInt32(c.Attribute("min").Value) == ColumnUtil.GetColumnIndex("F") && Convert.ToInt32(c.Attribute("max").Value) == ColumnUtil.GetColumnIndex("F")).FirstOrDefault();
                Assert.NotNull(colF);

                Assert.Equal(NewWidth / 2 + 2, Convert.ToDouble(colF.Attribute("width").Value, NumberFormatInfo.InvariantInfo));
                
                Assert.Equal(NewWidth / 2 + 2, worksheet.GetColumn("F").Width);
            }

            //Changing last column of range.
            using (var excelFile = ExcelFile.Open(path))
            {
                var worksheet = excelFile.OpenWorksheet(sheetName);
                worksheet.GetColumn("J").Width = NewWidth / 2 + 1;
            }

            using (var excelFile = ExcelFile.Open(path))
            {
                var worksheet = excelFile.OpenWorksheet(sheetName);

                var data = (XElement)reflection.GetValue(reflection.GetValue(worksheet, "data"), "data");
                var colJ = data.Descendants(data.GetDefaultNamespace() + "col").Where(c => Convert.ToInt32(c.Attribute("min").Value) == ColumnUtil.GetColumnIndex("J") && Convert.ToInt32(c.Attribute("max").Value) == ColumnUtil.GetColumnIndex("J")).FirstOrDefault();
                Assert.NotNull(colJ);

                Assert.Equal(NewWidth / 2 + 1, Convert.ToDouble(colJ.Attribute("width").Value, NumberFormatInfo.InvariantInfo));
                
                Assert.Equal(NewWidth / 2 + 1, worksheet.GetColumn("J").Width);
            }
        }
    }
}