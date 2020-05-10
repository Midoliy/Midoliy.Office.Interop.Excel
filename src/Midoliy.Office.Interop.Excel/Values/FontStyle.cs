using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Midoliy.Office.Interop
{
    [Flags]
    public enum FontStyle : int
    {
        None = 0,
        Bold = 1 << 0,
        Italic = 1 << 1,
        Shadow = 1 << 2,
        Strikethrough = 1 << 3,
        Subscript = 1 << 4,
        Superscript = 1 << 5,
        SingleUnderline = 1 << 6,
        DoubleUnderline = 1 << 7,
    }
}
