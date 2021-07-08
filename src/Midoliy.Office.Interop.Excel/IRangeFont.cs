using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace Midoliy.Office.Interop
{
    public interface IRangeFont
    {
        string Name { get; set; }
        double Size { get; set; }
        Color Color { get; set; }
        ThemeColor ThemeColor { get; set; }
        Tint Tint { get; set; }
        FontStyle Style { get; set; }
        bool Bold { get; set; }
        bool Italic { get; set; }
        bool Shadow { get; set; }
        bool OutlineFont { get; set; }
        bool Strikethrough { get; set; }
        bool Subscript { get; set; }
        bool Superscript { get; set; }
        Underline Underline { get; set; }
    }
}
