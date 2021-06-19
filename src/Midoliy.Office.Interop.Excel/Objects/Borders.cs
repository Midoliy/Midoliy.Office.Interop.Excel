using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using MsExcel = Microsoft.Office.Interop.Excel;

namespace Midoliy.Office.Interop.Objects
{
    internal readonly struct Borders : IBorders
    {
        public IBorder this[BordersIndex index]
        {
            get => new Border(_borders.Item[(MsExcel.XlBordersIndex)index]);
        }
        public Color Color
        {
            get => (Color)_borders.Color;
            set => _borders.Color = value;
        }

        public ThemeColor ThemeColor
        {
            get => (ThemeColor)_borders.ThemeColor;
            set => _borders.ThemeColor = value;
        }

        public Tint Tint
        {
            get => (Tint)((float)_borders.TintAndShade * 100.0f);
            set => _borders.TintAndShade = ((float)value / 100.0f);
        }

        public LineStyle LineStyle
        {
            get => (LineStyle)_borders.LineStyle;
            set => _borders.LineStyle = value;
        }

        public BorderWeight Weight
        {
            get => (BorderWeight)_borders.Weight;
            set => _borders.Weight = value;
        }

        public Borders(MsExcel.Borders borders)
        {
            _borders = borders;
        }

        private readonly MsExcel.Borders _borders;
    }
}
