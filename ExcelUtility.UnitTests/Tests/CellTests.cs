using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Xunit;
using System.Globalization;

namespace ExcelUtility.UnitTests.Tests
{
    public class CellTests : IDisposable
    {
        ExcelFile excelFile;

        public CellTests()
        {
            excelFile = ExcelFile.Open(@"D:/temp/DefaultWorksheet.xlsx");
        }

        [Fact]
        public void TestDoubleValue()
        {
            const double doubleValue = 3.78942147;

            IWorksheet worksheet = excelFile.OpenWorksheet("Paosdpoasdp");
            var cell = worksheet.GetCell("A4");
            cell.DoubleValue = doubleValue;
            var decendants = cell.Data.Descendants("v");
            Assert.NotNull(decendants);
            var descendant = decendants.FirstOrDefault();
            Assert.NotNull(descendant);
            var value =  descendant.Value;
            Assert.NotNull(value);
            var convertedValue = Convert.ToDouble(value, CultureInfo.InvariantCulture);
            Assert.Equal(doubleValue, convertedValue);
        }

        public void TestStringValue()
        {

        }

        public void Dispose()
        {
            excelFile.Dispose();
        }
    }
}
