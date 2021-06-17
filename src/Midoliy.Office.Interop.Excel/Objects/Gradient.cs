using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using MsExcel = Microsoft.Office.Interop.Excel;

namespace Midoliy.Office.Interop.Objects
{
    internal readonly struct Gradient : IGradient
    {
        public void Clear() => _gradient.ColorStops.Clear();

        public void Add(double start, double end, Color startColor, Color endColor)
        {
            if (start < 0.0 || 1.0 < start)
                throw new Exception("'start' は 0.0~1.0 の間で指定する.");

            if (end < 0.0 || 1.0 < end)
                throw new Exception("'end' は 0.0~1.0 の間で指定する.");

            Clear();
            _gradient.ColorStops.Add(start).Color = startColor;
            _gradient.ColorStops.Add(end).Color = endColor;
        }

        public int Degree
        {
            get => (int)_gradient.Degree;
            set
            {
                if (value < 0 || 360 < value)
                    throw new Exception("グラディーションの角度 'Degree' は 0~360° の間で指定する.");
                _gradient.Degree = value;
            }
        }

        public double Left
        {
            get => (double)_gradient.RectangleLeft;
            set
            {
                if (value < 0.0 || 1.0 < value)
                    throw new Exception("収束位置 'Left' は 0.0~1.0 の間で指定する.");
                _gradient.RectangleLeft = value;
            }
        }
        public double Right
        {
            get => (double)_gradient.RectangleRight;
            set
            {
                if (value < 0.0 || 1.0 < value)
                    throw new Exception("収束位置 'Right' は 0.0~1.0 の間で指定する.");
                _gradient.RectangleRight = value;
            }
        }
        public double Top
        {
            get => (double)_gradient.RectangleTop;
            set
            {
                if (value < 0.0 || 1.0 < value)
                    throw new Exception("収束位置 'Top' は 0.0~1.0 の間で指定する.");
                _gradient.RectangleTop = value;
            }
        }
        public double Bottom
        {
            get => (double)_gradient.RectangleBottom;
            set
            {
                if (value < 0.0 || 1.0 < value)
                    throw new Exception("収束位置 'Bottom' は 0.0~1.0 の間で指定する.");
                _gradient.RectangleBottom = value;
            }
        }

        internal Gradient(dynamic gradient)
        {
            _gradient = gradient;
        }

        private readonly dynamic _gradient;
    }
}
