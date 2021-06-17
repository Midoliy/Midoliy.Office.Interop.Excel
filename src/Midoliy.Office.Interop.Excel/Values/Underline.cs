using System;

namespace Midoliy.Office.Interop
{
    [Flags]
    public enum Underline : int
    {
        /// <summary>下線なし</summary>
        None = -4142,
        /// <summary>二重下線</summary>
        Double = -4119,
        /// <summary>二重下線(会計)</summary>
        DoubleAccounting = 5,
        /// <summary>一重下線</summary>
        Single = 2,
        /// <summary>一重下線（会計）</summary>
        SingleAccounting = 4,
    }
}
