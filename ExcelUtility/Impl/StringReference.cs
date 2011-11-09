using System.Xml.Linq;

namespace ExcelUtility.Impl
{
    public class StringReference
    {
        internal XElement Reference { get; set; }
        public int Index { get; set; }
        public int OldIndex { get; set; }
        public string Text { get; set; }
    }
}
