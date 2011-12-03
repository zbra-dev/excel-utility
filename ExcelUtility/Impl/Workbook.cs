using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using ExcelUtility.Utils;

namespace ExcelUtility.Impl
{
    internal class Workbook
    {
        private string rootFolder;
        private string path;
        private string xlFolder;
        private ContentTypes contentTypes;
        private List<IWorksheet> worksheets = new List<IWorksheet>();
        private XElementData data;
        private SharedStrings sharedStrings;

        public IEnumerable<IWorksheet> Worksheets { get { return worksheets; } }

        public Workbook(string rootFolder, ContentTypes contentTypes, SharedStrings sharedStrings)
        {
            this.rootFolder = rootFolder;
            this.contentTypes = contentTypes;
            this.sharedStrings = sharedStrings;
            this.path = rootFolder + contentTypes.GetWorkbookPath();
            this.xlFolder = Path.GetDirectoryName(path);
            ReadContents();
        }

        private void ReadContents()
        {
            data = new XElementData(XDocument.Load(path).Root);
            var relationshipsData = new XElementData(XDocument.Load(string.Format("{0}/_rels/workbook.xml.rels", Path.GetDirectoryName(path))).Root);
            foreach (var worksheetElement in data.Element("sheets").Descendants("sheet"))
            {
                var id = worksheetElement.AttributeValue("r", "id");
                var name = worksheetElement["name"];
                if (!id.StartsWith("rId"))
                    throw new InvalidDataException(string.Format("Invalid sheet id [{0}]", id));

                int sheetId = int.Parse(id.Substring("rId".Length), NumberFormatInfo.InvariantInfo);
                var target = relationshipsData.Descendants("Relationship").Single(r => r["Id"] == id)["Target"];
                var worksheetPath = string.Format("{0}/{1}", xlFolder, target);
                var worksheetFolder = Path.GetDirectoryName(worksheetPath);
                var worksheetData = new XElementData(XDocument.Load(worksheetPath).Root);
                worksheets.Add(new Worksheet(worksheetData, worksheetFolder, sharedStrings, name, sheetId));
            }
        }

        public void Save()
        {
            foreach (var worksheet in worksheets)
                worksheet.Save();
        }
    }
}
