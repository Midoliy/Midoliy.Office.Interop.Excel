using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace Midoliy.Office.Interop
{
    public interface IBorders
    {
        IBorder this[BordersIndex index] { get; }
        Color Color { get; set; }
        ThemeColor ThemeColor { get; set; }
        Tint Tint { get; set; }
        LineStyle LineStyle { get; set; }
        BorderWeight Weight { get; set; }
    }
}
