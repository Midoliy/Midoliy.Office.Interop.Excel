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
        }

        public static string ToColumnName(this int @this)
        {
            if (@this <= 0)
                throw new ArgumentException("Should be a value greater than 0.");

            var key = @this - 1;
            if (_columnNameStorage.TryGetValue(key, out string name))
                return name;

            var acc = new Stack<char>();
            while (0 < @this)
            {
                @this -= 1;
                var surplus = (@this % 26);
                acc.Push((char)((int)'A' + surplus));
                @this /= 26;
            }
            var col = new string(acc.ToArray());
            _columnNameStorage.Add(key, col);

            return col;
        }

        private static readonly Dictionary<int, string> _columnNameStorage;
    }
}
