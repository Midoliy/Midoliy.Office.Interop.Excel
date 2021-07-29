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
        /// <summary>100% 積み上げ面</summary>
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

    internal static class XlChartTypeEx
    {
        internal static int GetChartStyle(this ChartType type, ChartStyle style)
        {
            if (style == ChartStyle.Default)
                return -1;

            switch (type)
            {
                case ChartType.XlColumnClustered:
                    {
                        switch (style)
                        {
                            case ChartStyle.Style1: return 201;
                            case ChartStyle.Style2: return 202;
                            case ChartStyle.Style3: return 203;
                            case ChartStyle.Style4: return 204;
                            case ChartStyle.Style5: return 205;
                            case ChartStyle.Style6: return 206;
                            case ChartStyle.Style7: return 207;
                            case ChartStyle.Style8: return 208;
                            case ChartStyle.Style9: return 209;
                            case ChartStyle.Style10: return 210;
                            case ChartStyle.Style11: return 211;
                            case ChartStyle.Style12: return 212;
                            case ChartStyle.Style13: return 213;
                            case ChartStyle.Style14: return 214;
                            case ChartStyle.Style15: return 215;
                            case ChartStyle.Style16: return 340;
                            default: return -1;
                        }
                    }
                case ChartType.XlColumnStacked:
                case ChartType.XlColumnStacked100:
                    {
                        switch (style)
                        {
                            case ChartStyle.Style1: return 297;
                            case ChartStyle.Style2: return 298;
                            case ChartStyle.Style3: return 299;
                            case ChartStyle.Style4: return 300;
                            case ChartStyle.Style5: return 301;
                            case ChartStyle.Style6: return 302;
                            case ChartStyle.Style7: return 303;
                            case ChartStyle.Style8: return 304;
                            case ChartStyle.Style9: return 305;
                            case ChartStyle.Style10: return 306;
                            case ChartStyle.Style11: return 348;
                            default: return -1;
                        }
                    }
                case ChartType.Xl3DColumnClustered:
                    {
                        switch (style)
                        {
                            case ChartStyle.Style1: return 286;
                            case ChartStyle.Style2: return 287;
                            case ChartStyle.Style3: return 288;
                            case ChartStyle.Style4: return 289;
                            case ChartStyle.Style5: return 290;
                            case ChartStyle.Style6: return 291;
                            case ChartStyle.Style7: return 292;
                            case ChartStyle.Style8: return 293;
                            case ChartStyle.Style9: return 294;
                            case ChartStyle.Style10: return 295;
                            case ChartStyle.Style11: return 296;
                            case ChartStyle.Style12: return 347;
                            default: return -1;
                        }
                    }
                case ChartType.Xl3DColumnStacked:
                case ChartType.Xl3DColumnStacked100:
                    {
                        switch (style)
                        {
                            case ChartStyle.Style1: return 286;
                            case ChartStyle.Style2: return 299;
                            case ChartStyle.Style3: return 310;
                            case ChartStyle.Style4: return 289;
                            case ChartStyle.Style5: return 290;
                            case ChartStyle.Style6: return 294;
                            case ChartStyle.Style7: return 296;
                            case ChartStyle.Style8: return 347;
                            default: return -1;
                        }
                    }
                case ChartType.XlBarClustered:
                    {
                        switch (style)
                        {
                            case ChartStyle.Style1: return 216;
                            case ChartStyle.Style2: return 217;
                            case ChartStyle.Style3: return 218;
                            case ChartStyle.Style4: return 219;
                            case ChartStyle.Style5: return 220;
                            case ChartStyle.Style6: return 221;
                            case ChartStyle.Style7: return 222;
                            case ChartStyle.Style8: return 223;
                            case ChartStyle.Style9: return 224;
                            case ChartStyle.Style10: return 225;
                            case ChartStyle.Style11: return 339;
                            case ChartStyle.Style12: return 226;
                            case ChartStyle.Style13: return 341;
                            default: return -1;
                        }
                    }
                case ChartType.XlBarStacked:
                case ChartType.XlBarStacked100:
                    {
                        switch (style)
                        {
                            case ChartStyle.Style1: return 297;
                            case ChartStyle.Style2: return 298;
                            case ChartStyle.Style3: return 299;
                            case ChartStyle.Style4: return 300;
                            case ChartStyle.Style5: return 301;
                            case ChartStyle.Style6: return 302;
                            case ChartStyle.Style7: return 303;
                            case ChartStyle.Style8: return 304;
                            case ChartStyle.Style9: return 305;
                            case ChartStyle.Style10: return 306;
                            case ChartStyle.Style11: return 348;
                            default: return -1;
                        }
                    }
                case ChartType.Xl3DBarClustered:
                    {
                        switch (style)
                        {
                            case ChartStyle.Style1: return 286;
                            case ChartStyle.Style2: return 287;
                            case ChartStyle.Style3: return 288;
                            case ChartStyle.Style4: return 289;
                            case ChartStyle.Style5: return 290;
                            case ChartStyle.Style6: return 291;
                            case ChartStyle.Style7: return 292;
                            case ChartStyle.Style8: return 349;
                            case ChartStyle.Style9: return 294;
                            case ChartStyle.Style10: return 295;
                            case ChartStyle.Style11: return 296;
                            case ChartStyle.Style12: return 347;
                            default: return -1;
                        }
                    }
                case ChartType.Xl3DBarStacked:
                case ChartType.Xl3DBarStacked100:
                    {
                        switch (style)
                        {
                            case ChartStyle.Style1: return 286;
                            case ChartStyle.Style2: return 299;
                            case ChartStyle.Style3: return 310;
                            case ChartStyle.Style4: return 289;
                            case ChartStyle.Style5: return 290;
                            case ChartStyle.Style6: return 294;
                            case ChartStyle.Style7: return 296;
                            case ChartStyle.Style8: return 347;
                            default: return -1;
                        }
                    }
                case ChartType.XlLineStacked:
                case ChartType.XlLineStacked100:
                    {
                        switch (style)
                        {
                            case ChartStyle.Style1: return 227;
                            case ChartStyle.Style2: return 228;
                            case ChartStyle.Style3: return 229;
                            case ChartStyle.Style4: return 230;
                            case ChartStyle.Style5: return 231;
                            case ChartStyle.Style6: return 232;
                            case ChartStyle.Style7: return 233;
                            case ChartStyle.Style8: return 234;
                            case ChartStyle.Style9: return 235;
                            case ChartStyle.Style10: return 236;
                            case ChartStyle.Style11: return 237;
                            case ChartStyle.Style12: return 238;
                            case ChartStyle.Style13: return 239;
                            case ChartStyle.Style14: return 332;
                            case ChartStyle.Style15: return 342;
                            default: return -1;
                        }
                    }
                case ChartType.XlLineMarkers:
                case ChartType.XlLineMarkersStacked:
                    {
                        switch (style)
                        {
                            case ChartStyle.Style1: return 235;
                            case ChartStyle.Style2: return 236;
                            case ChartStyle.Style3: return 237;
                            case ChartStyle.Style4: return 238;
                            case ChartStyle.Style5: return 239;
                            case ChartStyle.Style6: return 332;
                            case ChartStyle.Style7: return 342;
                            case ChartStyle.Style8: return 227;
                            case ChartStyle.Style9: return 228;
                            case ChartStyle.Style10: return 229;
                            case ChartStyle.Style11: return 230;
                            case ChartStyle.Style12: return 231;
                            case ChartStyle.Style13: return 232;
                            case ChartStyle.Style14: return 233;
                            case ChartStyle.Style15: return 234;
                            default: return -1;
                        }
                    }
                case ChartType.XlLineMarkersStacked100:
                    {
                        switch (style)
                        {
                            case ChartStyle.Style1: return 227;
                            case ChartStyle.Style2: return 228;
                            case ChartStyle.Style3: return 229;
                            case ChartStyle.Style4: return 230;
                            case ChartStyle.Style5: return 231;
                            case ChartStyle.Style6: return 232;
                            case ChartStyle.Style7: return 233;
                            case ChartStyle.Style8: return 234;
                            case ChartStyle.Style9: return 235;
                            case ChartStyle.Style10: return 236;
                            case ChartStyle.Style11: return 237;
                            case ChartStyle.Style12: return 238;
                            case ChartStyle.Style13: return 239;
                            case ChartStyle.Style14: return 332;
                            case ChartStyle.Style15: return 342;
                            default: return -1;
                        }
                    }
                case ChartType.XlPieOfPie:
                    {
                        switch (style)
                        {
                            case ChartStyle.Style1: return 333;
                            case ChartStyle.Style2: return 252;
                            case ChartStyle.Style3: return 334;
                            case ChartStyle.Style4: return 335;
                            case ChartStyle.Style5: return 336;
                            case ChartStyle.Style6: return 337;
                            case ChartStyle.Style7: return 338;
                            case ChartStyle.Style8: return 258;
                            case ChartStyle.Style9: return 259;
                            case ChartStyle.Style10: return 260;
                            case ChartStyle.Style11: return 261;
                            default: return -1;
                        }
                    }
                case ChartType.XlPieExploded:
                    {
                        switch (style)
                        {
                            case ChartStyle.Style1: return 251;
                            case ChartStyle.Style2: return 252;
                            case ChartStyle.Style3: return 253;
                            case ChartStyle.Style4: return 254;
                            case ChartStyle.Style5: return 255;
                            case ChartStyle.Style6: return 256;
                            case ChartStyle.Style7: return 257;
                            case ChartStyle.Style8: return 258;
                            case ChartStyle.Style9: return 259;
                            case ChartStyle.Style10: return 260;
                            case ChartStyle.Style11: return 261;
                            default: return -1;
                        }
                    }
                case ChartType.Xl3DPieExploded:
                    {
                        switch (style)
                        {
                            case ChartStyle.Style1: return 262;
                            case ChartStyle.Style2: return 263;
                            case ChartStyle.Style3: return 264;
                            case ChartStyle.Style4: return 265;
                            case ChartStyle.Style5: return 266;
                            case ChartStyle.Style6: return 267;
                            case ChartStyle.Style7: return 268;
                            case ChartStyle.Style8: return 259;
                            case ChartStyle.Style9: return 261;
                            case ChartStyle.Style10: return 345;
                            default: return -1;
                        }
                    }
                case ChartType.XlBarOfPie:
                    {
                        switch (style)
                        {
                            case ChartStyle.Style1: return 333;
                            case ChartStyle.Style2: return 252;
                            case ChartStyle.Style3: return 334;
                            case ChartStyle.Style4: return 335;
                            case ChartStyle.Style5: return 336;
                            case ChartStyle.Style6: return 337;
                            case ChartStyle.Style7: return 338;
                            case ChartStyle.Style8: return 258;
                            case ChartStyle.Style9: return 259;
                            case ChartStyle.Style10: return 260;
                            case ChartStyle.Style11: return 261;
                            case ChartStyle.Style12: return 344;
                            default: return -1;
                        }
                    }
                case ChartType.XlXYScatterSmooth:
                case ChartType.XlXYScatterSmoothNoMarkers:
                case ChartType.XlXYScatterLines:
                case ChartType.XlXYScatterLinesNoMarkers:
                    {
                        switch (style)
                        {
                            case ChartStyle.Style1: return 240;
                            case ChartStyle.Style2: return 241;
                            case ChartStyle.Style3: return 242;
                            case ChartStyle.Style4: return 243;
                            case ChartStyle.Style5: return 244;
                            case ChartStyle.Style6: return 245;
                            case ChartStyle.Style7: return 246;
                            case ChartStyle.Style8: return 247;
                            case ChartStyle.Style9: return 248;
                            case ChartStyle.Style10: return 249;
                            case ChartStyle.Style11: return 250;
                            case ChartStyle.Style12: return 343;
                            default: return -1;
                        }
                    }
                case ChartType.XlAreaStacked:
                case ChartType.XlAreaStacked100:
                    {
                        switch (style)
                        {
                            case ChartStyle.Style1: return 276;
                            case ChartStyle.Style2: return 277;
                            case ChartStyle.Style3: return 278;
                            case ChartStyle.Style4: return 279;
                            case ChartStyle.Style5: return 280;
                            case ChartStyle.Style6: return 281;
                            case ChartStyle.Style7: return 282;
                            case ChartStyle.Style8: return 283;
                            case ChartStyle.Style9: return 284;
                            case ChartStyle.Style10: return 285;
                            case ChartStyle.Style11: return 346;
                            default: return -1;
                        }
                    }
                case ChartType.Xl3DAreaStacked:
                case ChartType.Xl3DAreaStacked100:
                    {
                        switch (style)
                        {
                            case ChartStyle.Style1: return 276;
                            case ChartStyle.Style2: return 312;
                            case ChartStyle.Style3: return 313;
                            case ChartStyle.Style4: return 314;
                            case ChartStyle.Style5: return 280;
                            case ChartStyle.Style6: return 281;
                            case ChartStyle.Style7: return 282;
                            case ChartStyle.Style8: return 315;
                            case ChartStyle.Style9: return 315;
                            case ChartStyle.Style10: return 350;
                            default: return -1;
                        }
                    }
                case ChartType.XlDoughnutExploded:
                    {
                        switch (style)
                        {
                            case ChartStyle.Style1: return 251;
                            case ChartStyle.Style2: return 252;
                            case ChartStyle.Style3: return 253;
                            case ChartStyle.Style4: return 254;
                            case ChartStyle.Style5: return 255;
                            case ChartStyle.Style6: return 256;
                            case ChartStyle.Style7: return 257;
                            case ChartStyle.Style8: return 258;
                            case ChartStyle.Style9: return 259;
                            case ChartStyle.Style10: return 260;
                            case ChartStyle.Style11: return 261;
                            default: return -1;
                        }
                    }
                case ChartType.XlRadar:
                case ChartType.XlRadarMarkers:
                case ChartType.XlRadarFilled:
                    {
                        switch (style)
                        {
                            case ChartStyle.Style1: return 317;
                            case ChartStyle.Style2: return 318;
                            case ChartStyle.Style3: return 206;
                            case ChartStyle.Style4: return 319;
                            case ChartStyle.Style5: return 320;
                            case ChartStyle.Style6: return 207;
                            case ChartStyle.Style7: return 321;
                            case ChartStyle.Style8: return 351;
                            default: return -1;
                        }
                    }
                case ChartType.XlSurface:
                case ChartType.XlSurfaceWireframe:
                case ChartType.XlSurfaceTopView:
                case ChartType.XlSurfaceTopViewWireframe:
                    {
                        switch (style)
                        {
                            case ChartStyle.Style1: return 307;
                            case ChartStyle.Style2: return 311;
                            case ChartStyle.Style3: return 308;
                            case ChartStyle.Style4: return 309;
                            default: return -1;
                        }
                    }
                case ChartType.XlBubble:
                case ChartType.XlBubble3DEffect:
                    {
                        switch (style)
                        {
                            case ChartStyle.Style1: return 269;
                            case ChartStyle.Style2: return 270;
                            case ChartStyle.Style3: return 271;
                            case ChartStyle.Style4: return 272;
                            case ChartStyle.Style5: return 246;
                            case ChartStyle.Style6: return 242;
                            case ChartStyle.Style7: return 273;
                            case ChartStyle.Style8: return 248;
                            case ChartStyle.Style9: return 274;
                            case ChartStyle.Style10: return 343;
                            default: return -1;
                        }
                    }
                case ChartType.XlStockHLC:
                case ChartType.XlStockOHLC:
                case ChartType.XlStockVHLC:
                case ChartType.XlStockVOHLC:
                    {
                        switch (style)
                        {
                            case ChartStyle.Style1: return 322;
                            case ChartStyle.Style2: return 323;
                            case ChartStyle.Style3: return 324;
                            case ChartStyle.Style4: return 325;
                            case ChartStyle.Style5: return 326;
                            case ChartStyle.Style6: return 327;
                            case ChartStyle.Style7: return 328;
                            case ChartStyle.Style8: return 329;
                            case ChartStyle.Style9: return 330;
                            case ChartStyle.Style10: return 331;
                            case ChartStyle.Style11: return 352;
                            default: return -1;
                        }
                    }
                case ChartType.XlCylinderColClustered:
                    {
                        switch (style)
                        {
                            case ChartStyle.Style1: return 286;
                            case ChartStyle.Style2: return 287;
                            case ChartStyle.Style3: return 288;
                            case ChartStyle.Style4: return 289;
                            case ChartStyle.Style5: return 290;
                            case ChartStyle.Style6: return 291;
                            case ChartStyle.Style7: return 292;
                            case ChartStyle.Style8: return 293;
                            case ChartStyle.Style9: return 294;
                            case ChartStyle.Style10: return 295;
                            case ChartStyle.Style11: return 296;
                            case ChartStyle.Style12: return 347;
                            default: return -1;
                        }
                    }
                case ChartType.XlCylinderColStacked:
                case ChartType.XlCylinderColStacked100:
                case ChartType.XlCylinderBarStacked:
                case ChartType.XlCylinderBarStacked100:
                case ChartType.XlConeColStacked:
                case ChartType.XlConeColStacked100:
                case ChartType.XlConeBarStacked:
                case ChartType.XlConeBarStacked100:
                case ChartType.XlPyramidColStacked:
                case ChartType.XlPyramidColStacked100:
                case ChartType.XlPyramidBarStacked:
                case ChartType.XlPyramidBarStacked100:
                    {
                        switch (style)
                        {
                            case ChartStyle.Style1: return 286;
                            case ChartStyle.Style2: return 299;
                            case ChartStyle.Style3: return 310;
                            case ChartStyle.Style4: return 289;
                            case ChartStyle.Style5: return 290;
                            case ChartStyle.Style6: return 294;
                            case ChartStyle.Style7: return 296;
                            case ChartStyle.Style8: return 347;
                            default: return -1;
                        }
                    }
                case ChartType.XlCylinderBarClustered:
                case ChartType.XlConeBarClustered:
                case ChartType.XlPyramidBarClustered:
                    {
                        switch (style)
                        {
                            case ChartStyle.Style1: return 286;
                            case ChartStyle.Style2: return 287;
                            case ChartStyle.Style3: return 288;
                            case ChartStyle.Style4: return 289;
                            case ChartStyle.Style5: return 290;
                            case ChartStyle.Style6: return 291;
                            case ChartStyle.Style7: return 292;
                            case ChartStyle.Style8: return 349;
                            case ChartStyle.Style9: return 294;
                            case ChartStyle.Style10: return 295;
                            case ChartStyle.Style11: return 296;
                            case ChartStyle.Style12: return 347;
                            default: return -1;
                        }
                    }
                case ChartType.XlCylinderCol:
                case ChartType.XlConeColClustered:
                case ChartType.XlConeCol:
                case ChartType.XlPyramidColClustered:
                    {
                        switch (style)
                        {
                            case ChartStyle.Style1: return 286;
                            case ChartStyle.Style2: return 287;
                            case ChartStyle.Style3: return 288;
                            case ChartStyle.Style4: return 289;
                            case ChartStyle.Style5: return 290;
                            case ChartStyle.Style6: return 291;
                            case ChartStyle.Style7: return 292;
                            case ChartStyle.Style8: return 293;
                            case ChartStyle.Style9: return 294;
                            case ChartStyle.Style10: return 295;
                            case ChartStyle.Style11: return 296;
                            case ChartStyle.Style12: return 347;
                            default: return -1;
                        }
                    }
                case ChartType.XlPyramidBarStacked100:
                    {
                        switch (style)
                        {
                            case ChartStyle.Style1: return 000;
                            case ChartStyle.Style2: return 000;
                            case ChartStyle.Style3: return 000;
                            case ChartStyle.Style4: return 000;
                            case ChartStyle.Style5: return 000;
                            case ChartStyle.Style6: return 000;
                            case ChartStyle.Style7: return 000;
                            case ChartStyle.Style8: return 000;
                            case ChartStyle.Style9: return 000;
                            case ChartStyle.Style10: return 000;
                            case ChartStyle.Style11: return 000;
                            case ChartStyle.Style12: return 000;
                            case ChartStyle.Style13: return 000;
                            case ChartStyle.Style14: return 000;
                            case ChartStyle.Style15: return 000;
                            default: return -1;
                        }
                    }
                case ChartType.XlPyramidCol: return -1;
                case ChartType.Xl3DColumn: return -1;
                case ChartType.XlLine: return 227;
                case ChartType.Xl3DLine: return -1;
                case ChartType.Xl3DPie: return -1;
                case ChartType.XlPie: return -1;
                case ChartType.XlXYScatter: return -1;
                case ChartType.Xl3DArea: return -1;
                case ChartType.XlArea: return -1;
                case ChartType.XlDoughnut: return -1;
                case ChartType.XlSunburst:
                    {
                        switch (style)
                        {
                            case ChartStyle.Style1: return 381;
                            case ChartStyle.Style2: return 382;
                            case ChartStyle.Style3: return 383;
                            case ChartStyle.Style4: return 384;
                            case ChartStyle.Style5: return 385;
                            case ChartStyle.Style6: return 386;
                            case ChartStyle.Style7: return 387;
                            case ChartStyle.Style8: return 388;
                            default: return -1;
                        }
                    }
                default: return -1;
            }
        }
    }
}
