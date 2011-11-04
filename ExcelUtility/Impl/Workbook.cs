using System.Collections.Generic;
using System.IO;

namespace ExcelUtility.Impl
{
    internal class Workbook
    {
        internal IList<IWorksheet> Worksheets { get; set; }
        internal string WorkbookPath { get; set; }

        internal Workbook()
        {
            Worksheets = new List<IWorksheet>();
        }

        internal void SaveChanges()
        {
            foreach (IWorksheet worksheet in Worksheets)
                worksheet.SaveChanges(string.Format("{0}worksheets/", Path.GetDirectoryName(WorkbookPath)));
        }
    }
}
