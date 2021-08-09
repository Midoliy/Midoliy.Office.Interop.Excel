using System;
using Midoliy.Office.Interop.Objects;

namespace Midoliy.Office.Interop
{
    public enum ChartType : int
    {
        /// <summary>散布図</summary>
        XlXYScatter = -4169,
        /// <summary>レーダー</summary>
        XlRadar = -4151,
        /// <summary>ドーナツ</summary>
        XlDoughnut = -4120,
        /// <summary>3-D 円</summary>
        Xl3DPie = -4102,
        /// <summary>3-D 折れ線</summary>
        Xl3DLine = -4101,
        /// <summary>3-D 縦棒</summary>
        Xl3DColumn = -4100,
        /// <summary>3-D 面</summary>
        Xl3DArea = -4098,
        /// <summary>面</summary>
        XlArea = 1,
        /// <summary>折れ線</summary>
        XlLine = 4,
        /// <summary>円</summary>
        XlPie = 5,
        /// <summary>バブル</summary>
        XlBubble = 15,
        /// <summary>集合縦棒</summary>
        XlColumnClustered = 51,
        /// <summary>積み上げ縦棒</summary>
        XlColumnStacked = 52,
        /// <summary>100% 積み上げ縦棒</summary>
        XlColumnStacked100 = 53,
        /// <summary>3-D 集合縦棒</summary>
        Xl3DColumnClustered = 54,
        /// <summary>3-D 積み上げ縦棒</summary>
        Xl3DColumnStacked = 55,
        /// <summary>3-D 100% 積み上げ縦棒</summary>
        Xl3DColumnStacked100 = 56,
        /// <summary>集合横棒</summary>
        XlBarClustered = 57,
        /// <summary>積み上げ横棒</summary>
        XlBarStacked = 58,
        /// <summary>100% 積み上げ横棒</summary>
        XlBarStacked100 = 59,
        /// <summary>3-D 集合横棒</summary>
        Xl3DBarClustered = 60,
        /// <summary>3-D 積み上げ横棒</summary>
        Xl3DBarStacked = 61,
        /// <summary>3-D 100% 積み上げ横棒</summary>
        Xl3DBarStacked100 = 62,
        /// <summary>積み上げ折れ線</summary>
        XlLineStacked = 63,
        /// <summary>100% 積み上げ折れ線</summary>
        XlLineStacked100 = 64,
        /// <summary>マーカー付き折れ線</summary>
        XlLineMarkers = 65,
        /// <summary>マーカー付き積み上げ折れ線</summary>
        XlLineMarkersStacked = 66,
        /// <summary>マーカー付き 100% 積み上げ折れ線</summary>
        XlLineMarkersStacked100 = 67,
        /// <summary>補助円グラフ付き円</summary>
        XlPieOfPie = 68,
        /// <summary>分割円</summary>
        XlPieExploded = 69,
        /// <summary>分割 3-D 円</summary>
        Xl3DPieExploded = 70,
        /// <summary>補助縦棒グラフ付き円</summary>
        XlBarOfPie = 71,
        /// <summary>平滑線付き散布図</summary>
        XlXYScatterSmooth = 72,
        /// <summary>平滑線付き散布図（データ マーカーなし）</summary>
        XlXYScatterSmoothNoMarkers = 73,
        /// <summary>折れ線付き散布図</summary>
        XlXYScatterLines = 74,
        /// <summary>折れ線付き散布図（データ マーカーなし）</summary>
        XlXYScatterLinesNoMarkers = 75,
        /// <summary>積み上げ面</summary>
        XlAreaStacked = 76,
        /// <summary>100% 積み上げ面</summary>
        XlAreaStacked100 = 77,
        /// <summary>3-D 積み上げ面</summary>
        Xl3DAreaStacked = 78,
        /// <summary>3-D 100% 積み上げ面</summary>
        Xl3DAreaStacked100 = 79,
        /// <summary>分割ドーナツ</summary>
        XlDoughnutExploded = 80,
        /// <summary>データ マーカー付きレーダー</summary>
        XlRadarMarkers = 81,
        /// <summary>塗りつぶしレーダー</summary>
        XlRadarFilled = 82,
        /// <summary>3-D 表面</summary>
        XlSurface = 83,
        /// <summary>3-D 表面（ワイヤーフレーム）</summary>
        XlSurfaceWireframe = 84,
        /// <summary>表面（トップビュー）</summary>
        XlSurfaceTopView = 85,
        /// <summary>表面（トップビュー・ワイヤーフレーム）</summary>
        XlSurfaceTopViewWireframe = 86,
        /// <summary>3-D 効果付きバブル</summary>
        XlBubble3DEffect = 87,
        /// <summary>高値-安値-終値</summary>
        XlStockHLC = 88,
        /// <summary>始値-高値-安値-終値</summary>
        XlStockOHLC = 89,
        /// <summary>出来高-高値-安値-終値</summary>
        XlStockVHLC = 90,
        /// <summary>出来高-始値-高値-安値-終値</summary>
        XlStockVOHLC = 91,
        /// <summary>集合円錐型 縦棒</summary>
        XlCylinderColClustered = 92,
        /// <summary>積み上げ円錐型 縦棒</summary>
        XlCylinderColStacked = 93,
        /// <summary>100% 積み上げ円柱型 縦棒</summary>
        XlCylinderColStacked100 = 94,
        /// <summary>集合円柱型 横棒</summary>
        XlCylinderBarClustered = 95,
        /// <summary>積み上げ円柱型 横棒</summary>
        XlCylinderBarStacked = 96,
        /// <summary>100% 積み上げ円柱型 横棒</summary>
        XlCylinderBarStacked100 = 97,
        /// <summary>3-D 円柱型 縦棒</summary>
        XlCylinderCol = 98,
        /// <summary>集合円錐型 縦棒</summary>
        XlConeColClustered = 99,
        /// <summary>積み上げ円錐型 縦棒</summary>
        XlConeColStacked = 100,
        /// <summary>100% 積み上げ円錐型 縦棒</summary>
        XlConeColStacked100 = 101,
        /// <summary>集合円錐型 横棒</summary>
        XlConeBarClustered = 102,
        /// <summary>積み上げ円錐型 横棒</summary>
        XlConeBarStacked = 103,
        /// <summary>100% 積み上げ円錐型 横棒</summary>
        XlConeBarStacked100 = 104,
        /// <summary>3-D 円錐型縦棒</summary>
        XlConeCol = 105,
        /// <summary>集合ピラミッド型 縦棒</summary>
        XlPyramidColClustered = 106,
        /// <summary>積み上げピラミッド型 縦棒</summary>
        XlPyramidColStacked = 107,
        /// <summary>100% 積み上げピラミッド型 縦棒</summary>
        XlPyramidColStacked100 = 108,
        /// <summary>集合ピラミッド型 横棒</summary>
        XlPyramidBarClustered = 109,
        /// <summary>積み上げピラミッド型 横棒</summary>
        XlPyramidBarStacked = 110,
        /// <summary>100% 積み上げピラミッド型 横棒</summary>
        XlPyramidBarStacked100 = 111,
        /// <summary>3-D ピラミッド型縦棒</summary>
        XlPyramidCol = 112,
        /// <summary>サンバースト</summary>
        XlSunburst = 120,
    }
}
