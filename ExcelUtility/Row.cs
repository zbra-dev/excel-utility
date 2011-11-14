
using System.Xml.Linq;
namespace ExcelUtility
{
    public class Row
    {
        private IWorksheet worksheet;
        private XElement rowData;
        private double height;

        public int Index { get; private set; }
        public double Height { get { return height; } set { SetHeight(value); } }
        public double Position { get { return worksheet.GetRowPosition(Index); } }

        internal Row(XElement rowData, IWorksheet worksheet, int index)
        {
            this.rowData = rowData;
            this.worksheet = worksheet;
            Index = index;
        }

        private void SetHeight(double value)
        {
            // change worksheet XML value
            height = value;
        }
    }
}
