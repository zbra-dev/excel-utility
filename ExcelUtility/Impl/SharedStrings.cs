using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using ExcelUtility.Utils;

namespace ExcelUtility.Impl
{
    internal class SharedStrings
    {
        private string path;
        private XElementData data;
        private Dictionary<string, SharedString> map = new Dictionary<string, SharedString>();

        public SharedStrings(string path)
        {
            this.path = path;
            data = new XElementData(XDocument.Load(path).Root);
            data.RemoveAttribute("count"); // optional values - will be recalculated
            data.RemoveAttribute("uniqueCount"); // optional values - will be recalculated
            ReadContents();
        }

        private void ReadContents()
        {
            foreach (var si in data.Descendants("si"))
            {
                var t = si.Element("t");
                if (t == null)
                    throw new ArgumentException("Invalid Shared Strings content");
                var sharedString = new SharedString() { Value = t.Value, Index = map.Count };
                map.Add(sharedString.Value, sharedString);
            }
        }

        public int GetStringReferenceOf(string value)
        {
            var normalizedValue = value ?? string.Empty;
            if (!map.TryGetValue(normalizedValue ?? string.Empty, out SharedString sharedString))
            {
                sharedString = new SharedString() { Value = normalizedValue, Index = map.Count };
                map.Add(normalizedValue, sharedString);
            }
            return sharedString.Index;
        }

        public void Save(IEnumerable<IWorksheet> worksheets)
        {
            CleanUpReferences(worksheets);
            data.Save(path);
        }

        private void CleanUpReferences(IEnumerable<IWorksheet> worksheets)
        {
            data.RemoveNodes();

            var cellMap = worksheets
                .SelectMany(w => w.DefinedRows)
                .SelectMany(r => r.DefinedCells)
                .Cast<Cell>()
                .Where(c => c.IsTypeString)
                .ToMultiMap(c => c.InternalValue);

            var list = map.Values.Where(s => cellMap.ContainsKey(s.Index.ToString())).ToList();
            for (int i = 0; i < list.Count; ++i)
            {
                data.Add("si").Add("t").Value = list[i].Value;
                int previousIndex = list[i].Index;
                foreach (var cell in cellMap[previousIndex.ToString()])
                    cell.InternalValue = i.ToString();
            }
        }

        public string GetValueByIndex(string index)
        {
            map.TryGetValue(index, out SharedString sharedString);
            return sharedString?.Value;
        }

        private class SharedString
        {
            public string Value { get; set; }
            public int Index { get; set; }
        }
    }
}
