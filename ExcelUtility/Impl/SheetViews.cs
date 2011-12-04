using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelUtility.Utils;

namespace ExcelUtility.Impl
{
    internal class SheetViews : ISheetViews
    {
        private XElementData data;
        private XElementData SheetView { get { return data.Element("sheetView"); } }

        public int TabSelected { get { return int.Parse(SheetView["tabSelected"]); } set { SheetView.SetAttributeValue("tabSelected", value); } }
        public int ZoomScale { get { return int.Parse(SheetView["zoomScale"]); } set { SheetView.SetAttributeValue("zoomScale", value); } }
        public int ZoomScaleNormal { get { return int.Parse(SheetView["zoomScaleNormal"]); } set { SheetView.SetAttributeValue("zoomScaleNormal", value); } }

        public SheetViews(XElementData data)
        {
            this.data = data;
        }
    }
}
