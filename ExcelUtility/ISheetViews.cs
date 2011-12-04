using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelUtility
{
    public interface ISheetViews
    {
        int TabSelected { get; set; }
        int ZoomScale { get; set; }
        int ZoomScaleNormal { get; set; }
    }
}
