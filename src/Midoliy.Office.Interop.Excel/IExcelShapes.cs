using System;
using System.Collections.Generic;

namespace Midoliy.Office.Interop
{
    public interface IExcelShapes
    {
        IExcelShape AddChart(ChartRecipe recipe, bool newLayout = true);
        IExcelShape AddChart(ChartRecipe recipe, IExcelRange range, bool newLayout);
        IExcelShape AddChart(ChartRecipe recipe, IExcelRange left, IExcelRange top, IExcelRange width, IExcelRange height, bool newLayout);
    }
}
