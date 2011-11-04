using System;
using System.Linq;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;
using ICSharpCode.SharpZipLib.Zip;
using System.Xml.Linq;
using ExcelUtility.Impl;

namespace ExcelUtility
{
    public class ExcelFile : IDisposable
    {
        private Workbook workbook;
        private string decompressPath;
        private string originalFilePath;

        public IList<IWorksheet> Worksheets { get { return workbook.Worksheets; } }

        // Find file and return a new instance of ExcelFile
        public static ExcelFile Open(string filePath)
        {
            return new ExcelFile(filePath);
        }

        private ExcelFile(string filePath)
        {
            ReadContent(filePath);
        }

        // Deflate xlsx and read Content_Types.xml and workbook.xml
        private void ReadContent(string filePath)
        {
            originalFilePath = filePath;
            //decompressPath = string.Format("{0}{1}/", Path.GetTempPath(), Path.GetFileNameWithoutExtension(filePath));
            decompressPath = string.Format(@"D:/temp/{0}/", Path.GetFileNameWithoutExtension(filePath));
            new FastZip().ExtractZip(filePath, decompressPath, null);
            workbook = GetWorkbook(decompressPath, XDocument.Load(string.Format("{0}[Content_Types].xml", decompressPath)));
            workbook.Worksheets = GetWorksheets(XDocument.Load(workbook.WorkbookPath));
        }

        private Workbook GetWorkbook(string rootPath, XDocument contentTypes)
        {
            return (from workbook in contentTypes.Descendants(XName.Get("Override", contentTypes.Root.GetDefaultNamespace().ToString()))
                    select new Workbook 
                    {
                        WorkbookPath = string.Format("{0}{1}", rootPath, workbook.Attribute("PartName").Value)
                    }).FirstOrDefault();
        }

        private IList<IWorksheet> GetWorksheets(XDocument workbook)
        {
            return (from sheet in workbook.Descendants(XName.Get("sheet", workbook.Root.GetDefaultNamespace().ToString()))
                    select (IWorksheet)new Worksheet()
                    {
                        Name = sheet.Attribute("name").Value,
                        SheetId = Convert.ToInt32(sheet.Attribute("sheetId").Value)
                    }).ToList();
        }

        // Regenarate xlsx file
        private void CloseFile()
        {
            SaveChanges();
            new FastZip().CreateZip(originalFilePath, decompressPath, true, null);
        }

        private void SaveChanges()
        {
            workbook.SaveChanges();
        }

        public IWorksheet OpenWorksheet(string name)
        {
            return workbook.Worksheets.FirstOrDefault<IWorksheet>(w => w.Name == name);
        }

        #region IDisposable Members

        public void Dispose()
        {
            CloseFile();
        }
        #endregion
    }
}
