using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelUtility
{
    public class ExcelFile : IDisposable
    {
        // Find file and return a new instance of ExcelFile
        public static ExcelFile Open(string filePath)
        {
            throw new NotImplementedException();
        }

        private ExcelFile(string filePath)
        {
            ReadContent(filePath);
        }

        // Deflate xlsx and read Content_Types.xml and workbook.xml
        private void ReadContent(string filePath)
        {
            throw new NotImplementedException();
        }

        // Regenarate xlsx file
        private void CloseFile()
        {
            throw new NotImplementedException();
        }

        public IWorksheet OpenWorksheet(string name)
        {
            throw new NotImplementedException();
        }

        #region IDisposable Members

        public void Dispose()
        {
            CloseFile();
        }

        #endregion

    }
}
