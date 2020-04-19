using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Midoliy.Office.Interop
{
    public static class IntegerEx
    {
        static IntegerEx()
        {
            _columnNameStorage = new Dictionary<int, string>();
            _columnNameTable = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            _tableSize = _columnNameTable.Length;
        }

        public static string  ToColumnName(this int @this)
        {
            var src = @this - 1;
            if (_columnNameStorage.TryGetValue(src, out string name))
                return name;

            var col = src.ToColumnNameStack().AsString();
            _columnNameStorage.Add(src, col);

            return col;
        }

        private static Stack<char> ToColumnNameStack(this int @this)
        {
            var acc = new Stack<char>();
            ToColumnName(@this, acc);
            return acc;
        }

        private static void ToColumnName(int i, Stack<char> acc)
        {
            if (i < 0)
                return;

            acc.Push(_columnNameTable[i % _tableSize]);

            if (_tableSize <= i)
                ToColumnName(i / _tableSize - 1, acc);
        }

        private static string AsString(this Stack<char> @this)
        {
            var acc = new StringBuilder();
            foreach (var c in @this)
                acc.Append(c);
            return acc.ToString();
        }

        private static readonly Dictionary<int, string> _columnNameStorage;
        private static readonly string _columnNameTable;
        private static readonly int _tableSize;
    }
}
