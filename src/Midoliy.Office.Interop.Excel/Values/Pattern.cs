using System;

namespace Midoliy.Office.Interop
{
    [Flags]
    public enum Pattern : int
    {
        /// <summary>縦 縞</summary>
        Vertical = -4166,
        /// <summary>左下がり斜線 縞</summary>
        Up = -4162,
        /// <summary>網掛け無し</summary>
        None = -4142,
        /// <summary>横 縞</summary>
        Horizontal = -4128,
        /// <summary>75% 灰色</summary>
        Gray75 = -4126,
        /// <summary>50% 灰色</summary>
        Gray50 = -4125,
        /// <summary>25% 灰色</summary>
        Gray25 = -4124,
        /// <summary>右下がり斜線 縞</summary>
        Down = -4121,
        /// <summary>自動設定</summary>
        Automatic = -4105,
        /// <summary>塗りつぶし</summary>
        Solid = 1,
        /// <summary>左下がり斜線 格子</summary>
        Checker = 9,
        /// <summary>極太線 左下がり斜線 格子</summary>
        SemiGray75 = 10,
        /// <summary>実線 横線</summary>
        LightHorizontal = 11,
        /// <summary>実線 縦 縞</summary>
        LightVertical = 12,
        /// <summary>実線 右下がり斜線 縞</summary>
        LightDown = 13,
        /// <summary>実線 左下がり斜線 縞</summary>
        LightUp = 14,
        /// <summary>実線 格子</summary>
        Grid = 15,
        /// <summary>実線 左下がりの斜線 格子</summary>
        CrissCross = 16,
        /// <summary>12.5% 灰色</summary>
        Gray16 = 17,
        /// <summary>6.25% 灰色</summary>
        Gray8 = 18,
        /// <summary>直線グラデーション. グラデーションの色・角度の指定が必要な場合有り.</summary>
        LinearGradient = 4000,
        /// <summary>角グラデーション. グラデーションの色・角度の指定が必要な場合有り.</summary>
        RectangularGradient = 4001,
    }
}
