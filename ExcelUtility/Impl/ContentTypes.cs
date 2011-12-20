using System.Linq;
using ExcelUtility.Utils;

namespace ExcelUtility.Impl
{
    public class ContentTypes
    {
        private const string WorkbookContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
        private const string WorksheetContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
        private const string StylesContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
        private const string SharedStringsContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
        private const string DrawingContentType = "application/vnd.openxmlformats-officedocument.drawing+xml";

        private XElementData data;
        private MultiMap<string, string> partNameMap = new MultiMap<string, string>();

        public ContentTypes(XElementData data)
        {
            this.data = data;
            ReadContents();
        }

        private void ReadContents()
        {
            foreach (var over in data.Descendants("Override"))
            {
                partNameMap.Add(over["ContentType"], over["PartName"]);
            }
        }

        public string GetWorkbookPath()
        {
            return partNameMap[WorkbookContentType].Single();
        }

        public string GetWorksheetPath(string sheetId)
        {
            return partNameMap[WorksheetContentType].Single(s => s.EndsWith(sheetId + ".xml"));
        }

        public string GetSharedStringsPath()
        {
            return partNameMap[SharedStringsContentType].Single();
        }
    }
}
