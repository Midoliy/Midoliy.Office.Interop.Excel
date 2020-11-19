using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Midoliy.Office.Interop
{
    public static class StringEx
    {
        static StringEx()
        {
            _columnNumberStorage = new Dictionary<string, int>();
            _columnNameTable = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        }

        public static int ToColumnNumber(this string @this, int depth = 1_000)
        {
            if (!@this.IsValidColumnName())
                throw new ArgumentException("The string must contain only A to Z.");
            if (depth < 1)
                throw new ArgumentException("'depth' must be greater than or equal to 1");

            if (_columnNumberStorage.TryGetValue(@this, out int key))
                return key;

            var found = IntegerEx.FindKey(@this);
            if (-1 < found)
            {
                found = found + 1;
                _columnNumberStorage.Add(@this, found);
                return found;
            }

            for (int i = 1; i < depth; i++)
            {
                if (@this == i.ToColumnName())
                {
                    _columnNumberStorage.Add(@this, i);
                    return i;
                }
            }
            return 0;
        }

        private static bool IsValidColumnName(this string @this)
        {
            foreach (var c in @this)
            {
                if (!_columnNameTable.Contains(c))
                    return false;
            }
            return true;
        }

        private static readonly Dictionary<string, int> _columnNumberStorage;
        private static readonly string _columnNameTable;
    }
}
