namespace Midoliy.Office.Interop
{
    public readonly struct ChartRecipe
    {
        public readonly ChartType Type;
        public readonly int Style;

        internal ChartRecipe(ChartType type, int style)
        {
            Type = type;
            Style = style;
        }

        public static ChartRecipe MakeColumnClustered(ColumnClusteredChartStyle style = ColumnClusteredChartStyle.Default) => new ChartRecipe(ChartType.XlColumnClustered, (int)style);
        public static ChartRecipe MakeColumnStacked(ColumnStackedChartStyle style = ColumnStackedChartStyle.Default) => new ChartRecipe(ChartType.XlColumnStacked, (int)style);
        public static ChartRecipe MakeColumnStacked100(ColumnStackedChartStyle style = ColumnStackedChartStyle.Default) => new ChartRecipe(ChartType.XlColumnStacked100, (int)style);
        public static ChartRecipe MakeBarStacked(BarStackedChartStyle style = BarStackedChartStyle.Default) => new ChartRecipe(ChartType.XlBarStacked, (int)style);
        public static ChartRecipe MakeBarStacked100(BarStackedChartStyle style = BarStackedChartStyle.Default) => new ChartRecipe(ChartType.XlBarStacked100, (int)style);
        public static ChartRecipe Make3DColumnClustered(Xl3DColumnClusteredChartStyle style = Xl3DColumnClusteredChartStyle.Default) => new ChartRecipe(ChartType.Xl3DColumnClustered, (int)style);
        public static ChartRecipe Make3DColumn(Xl3DColumnClusteredChartStyle style = Xl3DColumnClusteredChartStyle.Default) => new ChartRecipe(ChartType.Xl3DColumn, (int)style);
        public static ChartRecipe MakeConeColClustered(ConeColClusteredChartStyle style = ConeColClusteredChartStyle.Default) => new ChartRecipe(ChartType.XlConeColClustered, (int)style);
        public static ChartRecipe MakeConeCol(ConeColClusteredChartStyle style = ConeColClusteredChartStyle.Default) => new ChartRecipe(ChartType.XlConeCol, (int)style);
        public static ChartRecipe MakeCylinderColClustered(CylinderColClusteredChartStyle style = CylinderColClusteredChartStyle.Default) => new ChartRecipe(ChartType.XlCylinderColClustered, (int)style);
        public static ChartRecipe MakeCylinderCol(CylinderColClusteredChartStyle style = CylinderColClusteredChartStyle.Default) => new ChartRecipe(ChartType.XlCylinderCol, (int)style);
        public static ChartRecipe MakePyramidColClustered(PyramidColClusteredChartStyle style = PyramidColClusteredChartStyle.Default) => new ChartRecipe(ChartType.XlPyramidColClustered, (int)style);
        public static ChartRecipe MakePyramidCol(PyramidColClusteredChartStyle style = PyramidColClusteredChartStyle.Default) => new ChartRecipe(ChartType.XlPyramidCol, (int)style);
        public static ChartRecipe Make3DColumnStacked(Xl3DColumnStackedChartStyle style = Xl3DColumnStackedChartStyle.Default) => new ChartRecipe(ChartType.Xl3DColumnStacked, (int)style);
        public static ChartRecipe Make3DColumnStacked100(Xl3DColumnStackedChartStyle style = Xl3DColumnStackedChartStyle.Default) => new ChartRecipe(ChartType.Xl3DColumnStacked100, (int)style);
        public static ChartRecipe MakeCylinderColStacked(CylinderColStackedChartStyle style = CylinderColStackedChartStyle.Default) => new ChartRecipe(ChartType.XlCylinderColStacked, (int)style);
        public static ChartRecipe MakeCylinderColStacked100(CylinderColStackedChartStyle style = CylinderColStackedChartStyle.Default) => new ChartRecipe(ChartType.XlCylinderColStacked100, (int)style);
        public static ChartRecipe MakeCylinderBarStacked(CylinderBarStackedChartStyle style = CylinderBarStackedChartStyle.Default) => new ChartRecipe(ChartType.XlCylinderBarStacked, (int)style);
        public static ChartRecipe MakeCylinderBarStacked100(CylinderBarStackedChartStyle style = CylinderBarStackedChartStyle.Default) => new ChartRecipe(ChartType.XlCylinderBarStacked100, (int)style);
        public static ChartRecipe MakeConeColStacked(ConeColStackedChartStyle style = ConeColStackedChartStyle.Default) => new ChartRecipe(ChartType.XlConeColStacked, (int)style);
        public static ChartRecipe MakeConeColStacked100(ConeColStackedChartStyle style = ConeColStackedChartStyle.Default) => new ChartRecipe(ChartType.XlConeColStacked100, (int)style);
        public static ChartRecipe MakeConeBarStacked(ConeBarStackedChartStyle style = ConeBarStackedChartStyle.Default) => new ChartRecipe(ChartType.XlConeBarStacked, (int)style);
        public static ChartRecipe MakeConeBarStacked100(ConeBarStackedChartStyle style = ConeBarStackedChartStyle.Default) => new ChartRecipe(ChartType.XlConeBarStacked100, (int)style);
        public static ChartRecipe MakePyramidColStacked(PyramidColStackedChartStyle style = PyramidColStackedChartStyle.Default) => new ChartRecipe(ChartType.XlPyramidColStacked, (int)style);
        public static ChartRecipe MakePyramidColStacked100(PyramidColStackedChartStyle style = PyramidColStackedChartStyle.Default) => new ChartRecipe(ChartType.XlPyramidColStacked100, (int)style);
        public static ChartRecipe MakePyramidBarStacked(PyramidBarStackedChartStyle style = PyramidBarStackedChartStyle.Default) => new ChartRecipe(ChartType.XlPyramidBarStacked, (int)style);
        public static ChartRecipe MakePyramidBarStacked100(PyramidBarStackedChartStyle style = PyramidBarStackedChartStyle.Default) => new ChartRecipe(ChartType.XlPyramidBarStacked100, (int)style);
        public static ChartRecipe Make3DBarStacked(Xl3DBarStackedChartStyle style = Xl3DBarStackedChartStyle.Default) => new ChartRecipe(ChartType.Xl3DBarStacked, (int)style);
        public static ChartRecipe Make3DBarStacked100(Xl3DBarStackedChartStyle style = Xl3DBarStackedChartStyle.Default) => new ChartRecipe(ChartType.Xl3DBarStacked100, (int)style);
        public static ChartRecipe MakeBarClustered(BarClusteredChartStyle style = BarClusteredChartStyle.Default) => new ChartRecipe(ChartType.XlBarClustered, (int)style);
        public static ChartRecipe Make3DBarClustered(Xl3DBarClusteredChartStyle style = Xl3DBarClusteredChartStyle.Default) => new ChartRecipe(ChartType.Xl3DBarClustered, (int)style);
        public static ChartRecipe MakeCylinderBarClustered(CylinderBarClusteredChartStyle style = CylinderBarClusteredChartStyle.Default) => new ChartRecipe(ChartType.XlCylinderBarClustered, (int)style);
        public static ChartRecipe MakeConeBarClustered(ConeBarClusteredChartStyle style = ConeBarClusteredChartStyle.Default) => new ChartRecipe(ChartType.XlConeBarClustered, (int)style);
        public static ChartRecipe MakePyramidBarClustered(PyramidBarClusteredChartStyle style = PyramidBarClusteredChartStyle.Default) => new ChartRecipe(ChartType.XlPyramidBarClustered, (int)style);
        public static ChartRecipe MakeLineStacked(LineChartStyle style = LineChartStyle.Default) => new ChartRecipe(ChartType.XlLineStacked, (int)style);
        public static ChartRecipe MakeLineStacked100(LineChartStyle style = LineChartStyle.Default) => new ChartRecipe(ChartType.XlLineStacked100, (int)style);
        public static ChartRecipe MakeLineMarkersStacked100(LineChartStyle style = LineChartStyle.Default) => new ChartRecipe(ChartType.XlLineMarkersStacked100, (int)style);
        public static ChartRecipe MakeLine(LineChartStyle style = LineChartStyle.Default) => new ChartRecipe(ChartType.XlLine, (int)style);
        public static ChartRecipe MakeLineMarkers(LineMarkersChartStyle style = LineMarkersChartStyle.Default) => new ChartRecipe(ChartType.XlLineMarkers, (int)style);
        public static ChartRecipe MakeLineMarkersStacked(LineMarkersChartStyle style = LineMarkersChartStyle.Default) => new ChartRecipe(ChartType.XlLineMarkersStacked, (int)style);
        public static ChartRecipe MakePieOfPie(OfPieChartStyle style = OfPieChartStyle.Default) => new ChartRecipe(ChartType.XlPieOfPie, (int)style);
        public static ChartRecipe MakeBarOfPie(OfPieChartStyle style = OfPieChartStyle.Default) => new ChartRecipe(ChartType.XlBarOfPie, (int)style);
        public static ChartRecipe MakeDoughnut(DoughnutChartStyle style = DoughnutChartStyle.Default) => new ChartRecipe(ChartType.XlDoughnut, (int)style);
        public static ChartRecipe MakeDoughnutExploded(DoughnutChartStyle style = DoughnutChartStyle.Default) => new ChartRecipe(ChartType.XlDoughnutExploded, (int)style);
        public static ChartRecipe MakePie(PieChartStyle style = PieChartStyle.Default) => new ChartRecipe(ChartType.XlPie, (int)style);
        public static ChartRecipe MakePieExploded(PieChartStyle style = PieChartStyle.Default) => new ChartRecipe(ChartType.XlPieExploded, (int)style);
        public static ChartRecipe Make3DPieExploded(Xl3DPieChartStyle style = Xl3DPieChartStyle.Default) => new ChartRecipe(ChartType.Xl3DPieExploded, (int)style);
        public static ChartRecipe Make3DPie(Xl3DPieChartStyle style = Xl3DPieChartStyle.Default) => new ChartRecipe(ChartType.Xl3DPie, (int)style);
        public static ChartRecipe MakeXYScatterSmooth(XYScatterChartStyle style = XYScatterChartStyle.Default) => new ChartRecipe(ChartType.XlXYScatterSmooth, (int)style);
        public static ChartRecipe MakeXYScatterSmoothNoMarkers(XYScatterChartStyle style = XYScatterChartStyle.Default) => new ChartRecipe(ChartType.XlXYScatterSmoothNoMarkers, (int)style);
        public static ChartRecipe MakeXYScatterLines(XYScatterChartStyle style = XYScatterChartStyle.Default) => new ChartRecipe(ChartType.XlXYScatterLines, (int)style);
        public static ChartRecipe MakeXYScatterLinesNoMarkers(XYScatterChartStyle style = XYScatterChartStyle.Default) => new ChartRecipe(ChartType.XlXYScatterLinesNoMarkers, (int)style);
        public static ChartRecipe MakeXYScatter(XYScatterChartStyle style = XYScatterChartStyle.Default) => new ChartRecipe(ChartType.XlXYScatter, (int)style);
        public static ChartRecipe MakeAreaStacked(AreaChartStyle style = AreaChartStyle.Default) => new ChartRecipe(ChartType.XlAreaStacked, (int)style);
        public static ChartRecipe MakeAreaStacked100(AreaChartStyle style = AreaChartStyle.Default) => new ChartRecipe(ChartType.XlAreaStacked100, (int)style);
        public static ChartRecipe MakeArea(AreaChartStyle style = AreaChartStyle.Default) => new ChartRecipe(ChartType.XlArea, (int)style);
        public static ChartRecipe Make3DAreaStacked(Xl3DAreaChartStyle style = Xl3DAreaChartStyle.Default) => new ChartRecipe(ChartType.Xl3DAreaStacked, (int)style);
        public static ChartRecipe Make3DAreaStacked100(Xl3DAreaChartStyle style = Xl3DAreaChartStyle.Default) => new ChartRecipe(ChartType.Xl3DAreaStacked100, (int)style);
        public static ChartRecipe Make3DArea(Xl3DAreaChartStyle style = Xl3DAreaChartStyle.Default) => new ChartRecipe(ChartType.Xl3DArea, (int)style);
        public static ChartRecipe MakeRadar(RadarChartStyle style = RadarChartStyle.Default) => new ChartRecipe(ChartType.XlRadar, (int)style);
        public static ChartRecipe MakeRadarMarkers(RadarChartStyle style = RadarChartStyle.Default) => new ChartRecipe(ChartType.XlRadarMarkers, (int)style);
        public static ChartRecipe MakeRadarFilled(RadarChartStyle style = RadarChartStyle.Default) => new ChartRecipe(ChartType.XlRadarFilled, (int)style);
        public static ChartRecipe MakeSurface(SurfaceChartStyle style = SurfaceChartStyle.Default) => new ChartRecipe(ChartType.XlSurface, (int)style);
        public static ChartRecipe MakeSurfaceWireframe(SurfaceChartStyle style = SurfaceChartStyle.Default) => new ChartRecipe(ChartType.XlSurfaceWireframe, (int)style);
        public static ChartRecipe MakeSurfaceTopView(SurfaceChartStyle style = SurfaceChartStyle.Default) => new ChartRecipe(ChartType.XlSurfaceTopView, (int)style);
        public static ChartRecipe MakeSurfaceTopViewWireframe(SurfaceChartStyle style = SurfaceChartStyle.Default) => new ChartRecipe(ChartType.XlSurfaceTopViewWireframe, (int)style);
        public static ChartRecipe Make3DLine(Xl3DLineChartStyle style = Xl3DLineChartStyle.Default) => new ChartRecipe(ChartType.Xl3DLine, (int)style);
        public static ChartRecipe MakeBubble(BubbleChartStyle style = BubbleChartStyle.Default) => new ChartRecipe(ChartType.XlBubble, (int)style);
        public static ChartRecipe MakeBubble3DEffect(BubbleChartStyle style = BubbleChartStyle.Default) => new ChartRecipe(ChartType.XlBubble3DEffect, (int)style);
        public static ChartRecipe MakeStockHLC(StockChartStyle style = StockChartStyle.Default) => new ChartRecipe(ChartType.XlStockHLC, (int)style);
        public static ChartRecipe MakeStockOHLC(StockChartStyle style = StockChartStyle.Default) => new ChartRecipe(ChartType.XlStockOHLC, (int)style);
        public static ChartRecipe MakeStockVHLC(StockChartStyle style = StockChartStyle.Default) => new ChartRecipe(ChartType.XlStockVHLC, (int)style);
        public static ChartRecipe MakeStockVOHLC(StockChartStyle style = StockChartStyle.Default) => new ChartRecipe(ChartType.XlStockVOHLC, (int)style);
        public static ChartRecipe MakeSunburst(SunburstChartStyle style = SunburstChartStyle.Default) => new ChartRecipe(ChartType.XlSunburst, (int)style);
    }
}
