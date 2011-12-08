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

        public IEnumerable<IShape> Shapes { get { return map.Values.Cast<IShape>(); } }

        public Drawings(string path)
        {
            this.path = path;
            this.data = new XElementData("xdr", XDocument.Load(path).Root);
            map = data.Descendants("twoCellAnchor").Select(d => Shape.FromExisting(d)).ToDictionary(s => s.Id);
        }

        public IShape DrawShape(DrawPosition from, DrawPosition to)
        {
            var shape = Shape.New(data.Add("xdr", "twoCellAnchor"), map.Count == 0 ? 2 : map.Keys.Max() + 1, from, to);
            map.Add(shape.Id, shape);
            return shape;
        }

        public void Save()
        {
            data.Save(path);
        }

    }
}
