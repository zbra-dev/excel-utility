using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using ExcelUtility.Impl;
using ExcelUtility.Utils;
using ICSharpCode.SharpZipLib.Zip;

namespace ExcelUtility
{
    public class ExcelFile : IDisposable
    {
        public static ExcelFile Open(string filePath)
        {
            return new ExcelFile(filePath);
        }

        private ContentTypes contentTypes;
        private SharedStrings sharedStrings;
        private Workbook workbook;
        
        private string decompressFolder;
        private string filePath;

        private ExcelFile(string filePath)
        {
            ReadContents(filePath);
        }

        private void ReadContents(string filePath)
        {
            this.filePath = filePath;

            //decompressPath = string.Format("{0}{1}/", Path.GetTempPath(), Path.GetFileNameWithoutExtension(filePath));
            decompressFolder = string.Format(@"D:/Temp/{0}/", Path.GetFileNameWithoutExtension(filePath));
            new FastZip().ExtractZip(filePath, decompressFolder, null);

            var contentTypesData = new XElementData(XDocument.Load(string.Format("{0}[Content_Types].xml", decompressFolder)).Root);
            contentTypes = new ContentTypes(contentTypesData);
            sharedStrings = new SharedStrings(decompressFolder + contentTypes.GetSharedStringsPath());
            workbook = new Workbook(decompressFolder, contentTypes, sharedStrings);
        }

        private void Save()
        {
            sharedStrings.Save(workbook.Worksheets);
            workbook.Save();
            new FastZip().CreateZip(filePath, decompressFolder, true, null);
        }

        public IWorksheet OpenWorksheet(string name)
        {
            return workbook.Worksheets.FirstOrDefault(w => w.Name == name);
        }

        public void Dispose()
        {
            Save();
        }
    }
}