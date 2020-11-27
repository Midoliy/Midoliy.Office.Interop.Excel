using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Midoliy.Office.Interop
{
    public static class StringEx
    {
        static StringEx()
            => _columnNumberStorage = new Dictionary<string, int>();

        public static int ToColumnNumber(this string @this)
        {
            if (!@this.IsValidColumnName())
                throw new ArgumentOutOfRangeException("The string must contain only A to Z.");

            if (_columnNumberStorage.TryGetValue(@this, out int n))
                return n;

            var index = 0;
            foreach (int col in @this.ToUpper())
                index = (index * 26) + (col - 'A' + 1);

            _columnNumberStorage.Add(@this, index);
            return index;
        }

        private static bool IsValidColumnName(this string @this)
        {
            foreach (uint c in @this.ToUpper())
            {
                if (c < 'A' && 'Z' < c)
                    return false;
            }
            return true;
        }

        private static readonly Dictionary<string, int> _columnNumberStorage;
    }
}
