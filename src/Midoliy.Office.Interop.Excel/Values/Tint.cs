using System;

namespace Midoliy.Office.Interop
{
    [Flags]
    public enum Tint : int
    {
        /// <summary>明るさ -50%</summary>
        Dark50 = -50,
        /// <summary>明るさ -25%</summary>
        Dark25 = -25,
        /// <summary>中間色: 基準</summary>
        Default = 0,
        /// <summary>明るさ +40%</summary>
        Light40 = 40,
        /// <summary>明るさ +60%</summary>
        Light60 = 60,
        /// <summary>明るさ +80%</summary>
        Light80 = 80,
    }
}
