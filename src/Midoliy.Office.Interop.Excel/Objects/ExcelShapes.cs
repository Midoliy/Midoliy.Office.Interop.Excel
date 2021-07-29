using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Interop.Excel;
using MsExcel = Microsoft.Office.Interop.Excel;

namespace Midoliy.Office.Interop.Objects
{
    public enum ChartStyle
    {
        Default,
        Style1,
        Style2,
        Style3,
        Style4,
        Style5,
        Style6,
        Style7,
        Style8,
        Style9,
        Style10,
        Style11,
        Style12,
        Style13,
        Style14,
        Style15,
        Style16,
    }

    public readonly struct ChartRecipe
    {
        public readonly ChartType Type;
        public readonly ChartStyle Style;

        public ChartRecipe(ChartType type, ChartStyle style)
        {
            Type = type;
            Style = style;
        }
    }

    internal readonly struct ExcelShape : IExcelShape
    {
        internal ExcelShape(MsExcel.Shape shape)
        {
            _shape = shape;
        }

        private readonly MsExcel.Shape _shape;
    }

    internal readonly struct ExcelShapes : IExcelShapes
    {
        public IExcelShape AddChart(ChartRecipe chartType, IExcelRange range, bool newLayout) =>
            AddChart(chartType, range, range, range, range, newLayout);

        public IExcelShape AddChart(ChartRecipe chartType, IExcelRange left, IExcelRange top, IExcelRange width, IExcelRange height, bool newLayout)
        {
            var styleBase = chartType.Type.ToDefaultChartStyle();
            if (styleBase < 0)
                throw new ArgumentOutOfRangeException("");

            var style = (chartType.Style == ChartStyle.Default) ? -1 : styleBase + (int)chartType.Style;

            return new ExcelShape(_shapes.AddChart2(
                Style: style,
                XlChartType: chartType.Type,
                Left: left.Left,
                Top: top.Top,
                Width: width.Width,
                Height: height.Height,
                NewLayout: newLayout));
        }


        internal ExcelShapes(MsExcel.Shapes shapes)
        {
            _shapes = shapes;
            //_disposedValue = false;
        }

        private readonly MsExcel.Shapes _shapes;

        #region IDisposable Support
        //private bool _disposedValue;

        //private void Dispose(bool disposing)
        //{
        //    if (_disposedValue)
        //        return;

        //    if (disposing && _shapes != null)
        //    {
        //        try { while (0 < Marshal.ReleaseComObject(_shapes)) { } } catch { }
        //        _shapes = null;
        //        GC.Collect();
        //    }

        //    _disposedValue = true;
        //}

        //public void Dispose() => Dispose(true);
        #endregion


    }
}
