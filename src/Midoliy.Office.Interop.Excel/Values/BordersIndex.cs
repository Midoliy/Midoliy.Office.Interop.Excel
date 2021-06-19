using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Midoliy.Office.Interop
{
    public enum BordersIndex : int
    {
        /// <summary>範囲内の各セルの左上隅から右下に移動する罫線</summary>
        DiagonalDown = 5,
        /// <summary>範囲の各セルの左下隅から右上隅に移動する罫線</summary>
        DiagonalUp = 6,
        /// <summary>範囲内の下側の罫線</summary>
        EdgeBottom = 9,
        /// <summary>範囲の左側の罫線</summary>
        EdgeLeft = 7,
        /// <summary>範囲の右端の罫線</summary>
        EdgeRight = 10,
        /// <summary>範囲内の上側の罫線</summary>
        EdgeTop = 8,
        /// <summary>範囲外の罫線を除く、範囲内のすべてのセルの水平罫線</summary>
        InsideHorizontal = 12,
        /// <summary>範囲外の罫線を除く、範囲内のすべてのセルの垂直罫線</summary>
        InsideVertical = 11,
    }
}
