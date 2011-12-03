using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelUtility;
using Xunit;

namespace ExcelUtility.UnitTests.Tests
{
    public class WorksheetTests : IDisposable
    {
        private ExcelFile excelFile;

        public WorksheetTests()
        {
            excelFile = ExcelFile.Open(@"d:\temp\DefaultWorksheet.xlsx");
        }

        [Fact]
        public void OpenUnexistingWorksheet()
        {
            Assert.Throws<KeyNotFoundException>(() => { var sheet = excelFile.OpenWorksheet("aaa"); });
        }

        #region IDisposable Members

        public void Dispose()
        {
            excelFile.Dispose();
        }

        #endregion
    }
}
