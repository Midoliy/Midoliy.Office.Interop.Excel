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
    internal readonly struct ExcelShapes : IExcelShapes
    {
        public IExcelShape AddChart(ChartRecipe recipe, IExcelRange range, bool newLayout) =>
            AddChart(recipe, range, range, range, range, newLayout);

        public IExcelShape AddChart(ChartRecipe recipe, IExcelRange left, IExcelRange top, IExcelRange width, IExcelRange height, bool newLayout)
        {
            return new ExcelShape(_shapes.AddChart2(
                Style: recipe.Style,
                XlChartType: recipe.Type,
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
