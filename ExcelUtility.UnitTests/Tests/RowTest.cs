﻿using System;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using ExcelUtility.UnitTests.Util;
using Xunit;

namespace ExcelUtility.UnitTests.Tests
{
    public class RowTest
    {
        private const double NewHeight = 8;
        private const string path = @"d:\temp\DefaultWorksheet.xlsx";
        private const string sheetName = "Paosdpoasdp";
        private ReflectionUtil reflection;

        public RowTest()
        {
            reflection = new ReflectionUtil();
        }

        [Fact]
        public void GetRow()
        {
            //Ignored test with negative index.

            //Highest row.
            var excelFile = ExcelFile.Open(path);
            var worksheet = excelFile.OpenWorksheet(sheetName);

            var row = worksheet.GetRow(10255);
            Assert.NotNull(row);

            var data = (XElement)reflection.GetValue(reflection.GetValue(reflection.GetValue(worksheet, "sheetData"), "data"), "data");
            Assert.NotNull(data);

            var r = data.Descendants(data.GetDefaultNamespace() + "row").Where(t => t.Attribute("r").Value == "10255");
            Assert.NotNull(r.FirstOrDefault());
            excelFile.Save();
            excelFile.Close();

            //Lowest row.
            excelFile = ExcelFile.Open(path);
            worksheet = excelFile.OpenWorksheet(sheetName);
            row = worksheet.GetRow(1);
            Assert.NotNull(row);

            data = (XElement)reflection.GetValue(reflection.GetValue(reflection.GetValue(worksheet, "sheetData"), "data"), "data");
            Assert.NotNull(data);

            r = data.Descendants(data.GetDefaultNamespace() + "row").Where(t => t.Attribute("r").Value == "1");
            Assert.NotNull(r.FirstOrDefault());
            excelFile.Save();
            excelFile.Close();
        }

        [Fact]
        public void TestingRangeRows()
        {
            var excelFile = ExcelFile.Open(path);
            var worksheet = excelFile.OpenWorksheet(sheetName);
            worksheet.GetRow(10).Height = NewHeight;
            worksheet.GetRow(11).Height = NewHeight;
            worksheet.GetRow(12).Height = NewHeight;
            worksheet.GetRow(13).Height = NewHeight;
            worksheet.GetRow(14).Height = NewHeight;
            worksheet.GetRow(15).Height = NewHeight;
            worksheet.GetRow(16).Height = NewHeight;
            excelFile.Save();
            excelFile.Close();

            excelFile = ExcelFile.Open(path);
            worksheet = excelFile.OpenWorksheet(sheetName);

            var data = (XElement)reflection.GetValue(reflection.GetValue(worksheet, "data"), "data");

            var rows = data.Descendants(data.GetDefaultNamespace() + "row").Where(r => Convert.ToInt32(r.Attribute("r").Value) >= 10 && Convert.ToInt32(r.Attribute("r").Value) <= 16);
            Assert.NotNull(rows);

            var firstRow = rows.Where(r => Convert.ToInt32(r.Attribute("r")) == 10).FirstOrDefault();
            Assert.NotNull(firstRow);

            Assert.Equal(NewHeight, Convert.ToDouble(firstRow.Attribute("ht"), NumberFormatInfo.InvariantInfo));

            Assert.Equal(NewHeight, worksheet.GetRow(10).Height);
            Assert.Equal(NewHeight, worksheet.GetRow(11).Height);
            Assert.Equal(NewHeight, worksheet.GetRow(12).Height);
            Assert.Equal(NewHeight, worksheet.GetRow(13).Height);
            Assert.Equal(NewHeight, worksheet.GetRow(14).Height);
            Assert.Equal(NewHeight, worksheet.GetRow(15).Height);
            Assert.Equal(NewHeight, worksheet.GetRow(16).Height);
            excelFile.Save();
            excelFile.Close();

            //Change first row in a range.
            excelFile = ExcelFile.Open(path);
            worksheet = excelFile.OpenWorksheet(sheetName);
            worksheet.GetRow(23).Height = NewHeight;
            worksheet.GetRow(24).Height = NewHeight;
            worksheet.GetRow(25).Height = NewHeight;
            worksheet.GetRow(26).Height = NewHeight;
            excelFile.Save();
            excelFile.Close();

            excelFile = ExcelFile.Open(path);
            worksheet = excelFile.OpenWorksheet(sheetName);
            worksheet.GetRow(23).Height = NewHeight / 2;
            excelFile.Save();
            excelFile.Close();

            //Check
            excelFile = ExcelFile.Open(path);
            worksheet = excelFile.OpenWorksheet(sheetName);

            data = (XElement)reflection.GetValue(reflection.GetValue(worksheet, "data"), "data");

            var row = data.Descendants(data.GetDefaultNamespace() + "row").Where(r => Convert.ToInt32(r.Attribute("r").Value) == 23).FirstOrDefault();
            Assert.NotNull(row);

            Assert.Equal(NewHeight / 2, Convert.ToDouble(row.Attribute("ht"), NumberFormatInfo.InvariantInfo));

            Assert.Equal(NewHeight / 2, worksheet.GetRow(23).Height);
            excelFile.Save();
            excelFile.Close();

            //Change last row in a range
            excelFile = ExcelFile.Open(path);
            worksheet = excelFile.OpenWorksheet(sheetName);
            worksheet.GetRow(33).Height = NewHeight;
            worksheet.GetRow(34).Height = NewHeight;
            worksheet.GetRow(35).Height = NewHeight;
            worksheet.GetRow(36).Height = NewHeight;
            excelFile.Save();
            excelFile.Close();

            excelFile = ExcelFile.Open(path);
            worksheet = excelFile.OpenWorksheet(sheetName);
            worksheet.GetRow(36).Height = NewHeight / 2;
            excelFile.Save();
            excelFile.Close();

            excelFile = ExcelFile.Open(path);
            worksheet = excelFile.OpenWorksheet(sheetName);

            data = (XElement)reflection.GetValue(reflection.GetValue(worksheet, "data"), "data");

            row = data.Descendants(data.GetDefaultNamespace() + "row").Where(r => Convert.ToInt32(r.Attribute("r").Value) == 36).FirstOrDefault();
            Assert.NotNull(row);

            Assert.Equal(NewHeight / 2, Convert.ToDouble(row.Attribute("ht"), NumberFormatInfo.InvariantInfo));

            Assert.Equal(NewHeight / 2, worksheet.GetRow(36).Height);
            excelFile.Save();
            excelFile.Close();

            //Changing a midle row in the range
            excelFile = ExcelFile.Open(path);
            worksheet = excelFile.OpenWorksheet(sheetName);
            worksheet.GetRow(43).Height = NewHeight;
            worksheet.GetRow(44).Height = NewHeight;
            worksheet.GetRow(45).Height = NewHeight;
            worksheet.GetRow(46).Height = NewHeight;
            excelFile.Save();
            excelFile.Close();

            excelFile = ExcelFile.Open(path);
            worksheet = excelFile.OpenWorksheet(sheetName);
            worksheet.GetRow(44).Height = NewHeight / 2;
            excelFile.Save();
            excelFile.Close();

            excelFile = ExcelFile.Open(path);
            worksheet = excelFile.OpenWorksheet(sheetName);
            data = (XElement)reflection.GetValue(reflection.GetValue(worksheet, "data"), "data");

            row = data.Descendants(data.GetDefaultNamespace() + "row").Where(r => Convert.ToInt32(r.Attribute("r").Value) == 44).FirstOrDefault();
            Assert.NotNull(row);

            Assert.Equal(NewHeight / 2, Convert.ToDouble(row.Attribute("ht"), NumberFormatInfo.InvariantInfo));

            Assert.Equal(NewHeight / 2, worksheet.GetRow(44).Height);
            excelFile.Save();
            excelFile.Close();
        }
    }
}