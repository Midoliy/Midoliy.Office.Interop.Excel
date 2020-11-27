using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Midoliy.Office.Interop
{
    public static class StringEx
    {
        public static int ToColumnNumber(this string @this)
        {
            if (!@this.IsValidColumnName())
                throw new ArgumentException("The string must contain only A to Z.");

            var index = 0u;
            foreach (uint col in @this.ToUpper())
                index = index * 26 + ((uint)col - (uint)'A' + 1);
            return (int)index;
        }

        private static bool IsValidColumnName(this string @this)
        {
            foreach (uint c in @this.ToUpper())
            {
                if (c < (uint)'A' && (uint)'Z' < c)
                    return false;
            }
            return true;
        }
    }
}
