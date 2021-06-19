using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Midoliy.Office.Interop
{
    public enum LineStyle : int
    {
        /// <summary>線なし</summary>
        None = -4142,
        /// <summary>点線</summary>
        Dot = -4118,
        /// <summary>二重線</summary>
        Double = -4119,
        /// <summary>破線</summary>
        Dash = -4115,
        /// <summary>実線</summary>
        Continuous = 1,
        /// <summary>一点鎖線</summary>
        DashDot = 4,
        /// <summary>ニ点鎖線</summary>
        DashDotDot = 5,
        /// <summary>斜破線</summary>
        SlantDashDot = 13
    }
}
