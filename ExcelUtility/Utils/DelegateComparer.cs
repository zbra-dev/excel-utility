using System;
using System.Collections.Generic;

namespace ExcelUtility.Utils
{
    public class DelegateComparer<T> : IComparer<T>
    {
        private Comparison<T> comparison;

        public DelegateComparer(Comparison<T> comparison)
        {
            this.comparison = comparison;
        }

        #region IComparer<T> Members

        public int Compare(T x, T y)
        {
            return comparison(x, y);
        }

        #endregion
    }
}
