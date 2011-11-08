using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using ExcelUtility.Impl;
using ICSharpCode.SharpZipLib.Zip;

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
            workbook.Worksheets = GetWorksheets(workbook);
        }

        private Workbook GetWorkbook(string rootPath, XDocument contentTypes)
        {
            return (from workbook in contentTypes.Descendants(contentTypes.Root.GetDefaultNamespace() + "Override")
                    select new Workbook 
                    {
                        WorkbookPath = string.Format("{0}{1}", rootPath, workbook.Attribute("PartName").Value)
                    }).FirstOrDefault();
        }

        private IList<IWorksheet> GetWorksheets(Workbook workbook)
        {
            XDocument workbookData = XDocument.Load(workbook.WorkbookPath);
            return (from sheet in workbookData.Descendants(workbookData.Root.GetDefaultNamespace() + "sheet")
                    select (IWorksheet)new Worksheet(XElement.Load(string.Format("{0}/worksheets/sheet{1}.xml", Path.GetDirectoryName(workbook.WorkbookPath), sheet.Attribute("sheetId").Value)))
                    {
                        SharedStrings = new SharedStrings(XElement.Load(string.Format("{0}/sharedStrings.xml", Path.GetDirectoryName(workbook.WorkbookPath)))),
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
