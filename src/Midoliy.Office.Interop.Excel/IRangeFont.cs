using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Midoliy.Office.Interop
{
    public interface IRangeFont
    {
        double Size { get; set; }
        FontStyle Style { get; set; }
    }
}
