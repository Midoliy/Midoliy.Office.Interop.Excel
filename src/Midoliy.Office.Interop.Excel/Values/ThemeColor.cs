using System;

namespace Midoliy.Office.Interop
{
    [Flags]
    public enum ThemeColor : int
    {
        /// <summary>背景1</summary>
        Background1 = 1,
        /// <summary>テキスト1</summary>
        Foreground1 = 2,
        /// <summary>背景2</summary>
        Background2 = 3,
        /// <summary>テキスト2</summary>
        Foreground2 = 4,
        /// <summary>アクセント1</summary>
        Accent1 = 5,
        /// <summary>アクセント2</summary>
        Accent2 = 6,
        /// <summary>アクセント3</summary>
        Accent3 = 7,
        /// <summary>アクセント4</summary>
        Accent4 = 8,
        /// <summary>アクセント5</summary>
        Accent5 = 9,
        /// <summary>アクセント6</summary>
        Accent6 = 10,
    }
}
