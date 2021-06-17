using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace Midoliy.Office.Interop
{
    public interface IGradient
    {
        void Clear();
        void Add(double start, double end, Color startColor, Color endColor);
        int Degree { get; set; }
        double Left { get; set; }
        double Right { get; set; }
        double Top { get; set; }
        double Bottom { get; set; }
    }
}
