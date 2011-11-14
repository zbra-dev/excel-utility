using System.Collections.Generic;
using System.IO;

namespace ExcelUtility.Impl
{
    internal class Workbook
    {
        internal IList<IWorksheet> Worksheets { get; set; }
        internal string WorkbookPath { get; set; }

        public static string SharedStringsRelationshipType { get { return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"; } }
        public static string WorksheetRelationshipType { get { return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"; } }

        internal Workbook()
        {
            Worksheets = new List<IWorksheet>();
        }

        internal void SaveChanges()
        {
            foreach (IWorksheet worksheet in Worksheets)
                worksheet.SaveChanges(Path.GetDirectoryName(WorkbookPath));
        }
    }
}
