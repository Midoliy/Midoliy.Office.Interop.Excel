using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using MsExcel = Microsoft.Office.Interop.Excel;

namespace Midoliy.Office.Interop.Objects
{
    internal readonly struct Interior : IInterior
    {
        public Color Color 
        {
            get => (Color)_interior.Color;
            set => _interior.Color = value;
        }

        public Pattern Pattern
        {
            get => (Pattern)_interior.Pattern;
            set => _interior.Pattern = value;
        }

        public bool InvertIfNegative
        {
            get => (bool)_interior.InvertIfNegative;
            set => _interior.InvertIfNegative = value;
        }

        public ThemeColor ThemeColor
        {
            get => (ThemeColor)_interior.ThemeColor;
            set => _interior.ThemeColor = value;
        }

        public Tint Tint
        {
            get => (Tint)((float)_interior.TintAndShade * 100.0f);
            set => _interior.TintAndShade = ((float)value / 100.0f);
        }

        public IGradient Gradient => new Gradient(_interior.Gradient);

        public Interior(MsExcel.Interior interior)
        {
            _interior = interior;
        }

        private readonly MsExcel.Interior _interior;
    }
}
