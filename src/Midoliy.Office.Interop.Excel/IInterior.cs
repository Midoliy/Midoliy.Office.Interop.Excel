using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace Midoliy.Office.Interop
{
    public interface IInterior
    {
        Color Color { get; set; }
        Pattern Pattern { get; set; }
        bool InvertIfNegative { get; set; }
        ThemeColor ThemeColor { get; set; }
        Tint Tint { get; set; }
    }
}
