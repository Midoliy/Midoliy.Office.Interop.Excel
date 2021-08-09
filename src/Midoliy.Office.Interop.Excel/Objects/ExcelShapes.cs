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
        public IExcelShape AddChart(ChartRecipe recipe, bool newLayout) =>
            new ExcelShape(_shapes.AddChart2(Style: recipe.Style, XlChartType: recipe.Type, NewLayout: newLayout));
        public IExcelShape AddChart(ChartRecipe recipe, IExcelRange topLeft, bool newLayout) =>
            new ExcelShape( _shapes.AddChart2(Style: recipe.Style, XlChartType: recipe.Type, Left: topLeft.Left, Top: topLeft.Top, NewLayout: newLayout));
        public IExcelShape AddChart(ChartRecipe recipe, IExcelRange width, IExcelRange height, bool newLayout) =>
            new ExcelShape(_shapes.AddChart2(Style: recipe.Style, XlChartType: recipe.Type, Width: width.Width, Height: height.Height, NewLayout: newLayout));
        public IExcelShape AddChart(ChartRecipe recipe, IExcelRange topLeft, IExcelRange width, IExcelRange height, bool newLayout) =>
            new ExcelShape(_shapes.AddChart2(Style: recipe.Style, XlChartType: recipe.Type, Left: topLeft.Left, Top: topLeft.Top, Width: width.Width, Height: height.Height, NewLayout: newLayout));
        public IExcelShape AddChart(ChartRecipe recipe, IExcelRange left, IExcelRange top, IExcelRange width, IExcelRange height, bool newLayout) =>
            new ExcelShape(_shapes.AddChart2(Style: recipe.Style, XlChartType: recipe.Type, Left: left.Left, Top: top.Top, Width: width.Width, Height: height.Height, NewLayout: newLayout));

        internal ExcelShapes(MsExcel.Shapes shapes)
        {
            _shapes = shapes;
        }

        internal readonly MsExcel.Shapes _shapes;
    }
}
