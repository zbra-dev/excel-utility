using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using ExcelUtility.Utils;

namespace ExcelUtility.Impl
{
    internal class Drawings
    {
        private string path;
        private XElementData data;
        private Dictionary<int, Shape> map = new Dictionary<int, Shape>();
        
        public Drawings(string path)
        {
            this.path = path;
            this.data = new XElementData("xdr", XDocument.Load(path).Root);
            map = data.Descendants("twoCellAnchor").Select(d => Shape.FromExisting(d)).ToDictionary(s => s.Id);
        }

        public IShape DrawShape(DrawPosition from, DrawPosition to)
        {
            var shape = Shape.New(data.Add("xdr", "twoCellAnchor"), map.Count == 0 ? 2 : map.Keys.Max() + 1, from, to);
            return shape;
        }

        public void Save()
        {
            data.Save(path);
        }
    }
}
