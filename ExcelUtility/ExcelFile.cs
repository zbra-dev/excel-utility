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
        private string filePath;

        public string DecompressFolder { get; private set; }

        private ExcelFile(string filePath)
        {
            ReadContents(filePath);
        }

        private void ReadContents(string filePath)
        {
            this.filePath = filePath;

            DecompressFolder = Path.GetTempFileName() + @"_excelutility\";
            Directory.CreateDirectory(DecompressFolder);
            new FastZip().ExtractZip(filePath, DecompressFolder, null);

            var contentTypesData = new XElementData(XDocument.Load(string.Format("{0}[Content_Types].xml", DecompressFolder)).Root);
            contentTypes = new ContentTypes(contentTypesData);
            sharedStrings = new SharedStrings(DecompressFolder + contentTypes.GetSharedStringsPath());
            workbook = new Workbook(DecompressFolder, contentTypes, sharedStrings);
        }

        public void Save()
        {
            sharedStrings.Save(workbook.Worksheets);
            workbook.Save();
        }

        private void Close()
        {
            new FastZip().CreateZip(filePath, DecompressFolder, true, null);
            try
            {
                Directory.Delete(DecompressFolder, true);
            }
            catch (Exception)
            {
                // if can't delete temp folder then just log and ignore exception
            }
        }

        public IWorksheet OpenWorksheet(string name)
        {
            return workbook.Worksheets.FirstOrDefault(w => w.Name == name);
        }

        public void Dispose()
        {
            Close();
        }
    }
}