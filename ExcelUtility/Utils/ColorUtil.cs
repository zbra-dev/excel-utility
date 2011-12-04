using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelUtility.Utils
{
    public static class ColorUtil
    {
        public static int EncodeColor(int r, int g, int b)
        {
            int rgb = r;
            rgb = (r << 8) + g;
            rgb = (r << 8) + b;
            return rgb;
        }

        public static void DecodeColor(int rgb, out int r, out int g, out int b)
        {
            r = (rgb >> 16) & 0xFF;
            g = (rgb >> 8) & 0xFF;
            b = rgb & 0xFF;
        }
    }
}
