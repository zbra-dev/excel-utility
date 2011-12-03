using System;
using System.Linq;

namespace ExcelUtility.Utils
{
    public static class ColumnUtil
    {
        public static string GetColumnName(long index)
        {
            if (index < 0)
                throw new ArgumentException("Index can't be < 0", "index");
            string name = "";
            long div = index;
            long rest;
            bool first = true;
            do
            {
                div = Math.DivRem(div, 26, out rest);
                if (!first)
                    rest--;
                name = ((char)(rest + 'A')).ToString() + name;
                first = false;
            } while (div > 0);
            return name;
        }

        public static long GetColumnIndex(string name)
        {
            if (name.Length == 0)
                throw new ArgumentException("Invalid name", "name");
            long index = 0;
            int count = 0;
            bool first = true;
            foreach (char c in name.Reverse())
            {
                int diff = c - 'A';
                if (!first)
                    diff++;
                index += (long)Math.Pow(26, count++) * diff;
                first = false;
            }
            return index;
        }
    }
}
