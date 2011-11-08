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
        private SharedStrings sharedString;
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

            XDocument contentTypes = XDocument.Load(string.Format("{0}[Content_Types].xml", decompressPath));
            sharedString = BuildSharedString(decompressPath, contentTypes);
            workbook = BuildWorkbook(decompressPath, contentTypes);
            workbook.Worksheets = BuildWorksheets(workbook, sharedString);
        }

        private SharedStrings BuildSharedString(string rootPath, XDocument contentTypes)
        {
            return (from content in contentTypes.Descendants(contentTypes.Root.GetDefaultNamespace() + "Override")
                        where content.Attribute("PartName").Value.Contains("sharedStrings.xml")
                        select new SharedStrings(XElement.Load(string.Format("{0}/{1}", rootPath, content.Attribute("PartName").Value)))).FirstOrDefault();
        }

        private Workbook BuildWorkbook(string rootPath, XDocument contentTypes)
        {
            return (from content in contentTypes.Descendants(contentTypes.Root.GetDefaultNamespace() + "Override")
                    where content.Attribute("PartName").Value.Contains("workbook.xml")
                    select new Workbook
                    {
                        WorkbookPath = string.Format("{0}{1}", rootPath, content.Attribute("PartName").Value)
                    }).FirstOrDefault();
        }

        private IList<IWorksheet> BuildWorksheets(Workbook workbook, SharedStrings sharedStrings)
        {
            XDocument workbookData = XDocument.Load(workbook.WorkbookPath);
            return (from sheet in workbookData.Descendants(workbookData.Root.GetDefaultNamespace() + "sheet")
                    select (IWorksheet)new Worksheet(XElement.Load(string.Format("{0}/worksheets/sheet{1}.xml", Path.GetDirectoryName(workbook.WorkbookPath), sheet.Attribute("sheetId").Value)))
                    {
                        SharedStrings = sharedStrings,
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
            sharedString.CleanUpReferences(workbook.Worksheets);
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
