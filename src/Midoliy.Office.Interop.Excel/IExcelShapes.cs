using System;
using System.Collections.Generic;

namespace Midoliy.Office.Interop
{
    public interface IExcelShape// : IDisposable
    {

    }

    public interface IExcelShapes// : IDisposable
    {
        IExcelShape AddChart(ChartRecipe recipe);
        IExcelShape AddChart(ChartRecipe recipe, IExcelRange range, bool newLayout);
        IExcelShape AddChart(ChartRecipe recipe, IExcelRange left, IExcelRange top, IExcelRange width, IExcelRange height, bool newLayout);
    }
}
