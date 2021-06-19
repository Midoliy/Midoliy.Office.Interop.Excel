using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using MsExcel = Microsoft.Office.Interop.Excel;

namespace Midoliy.Office.Interop.Objects
{
    internal readonly struct Border : IBorder
    {
        public Color Color
        {
            get => (Color)_border.Color;
            set => _border.Color = value;
        }

        public ThemeColor ThemeColor
        {
            get => (ThemeColor)_border.ThemeColor;
            set => _border.ThemeColor = value;
        }

        public Tint Tint
        {
            get => (Tint)((float)_border.TintAndShade * 100.0f);
            set => _border.TintAndShade = ((float)value / 100.0f);
        }

        public LineStyle LineStyle
        {
            get => (LineStyle)_border.LineStyle;
            set => _border.LineStyle = value;
        }

        public BorderWeight Weight
        {
            get => (BorderWeight)_border.Weight;
            set => _border.Weight = value;
        }
        public Border(MsExcel.Border border)
        {
            _border = border;
        }

        private readonly MsExcel.Border _border;
    }
}
