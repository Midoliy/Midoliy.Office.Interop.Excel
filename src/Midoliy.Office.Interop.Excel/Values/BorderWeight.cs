using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Midoliy.Office.Interop
{
    public enum BorderWeight : int
    {
        /// <summary>普通</summary>
        Medium = -4138,
        /// <summary>細線 (最も細い罫線)</summary>
        Hairline = 1,
        /// <summary>極細</summary>
        Thin = 2,
        /// <summary>太線 (最も太い罫線)</summary>
        Thick = 4,
    }
}
