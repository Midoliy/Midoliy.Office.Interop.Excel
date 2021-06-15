using Midoliy.Office.Interop;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using MsExcel = Microsoft.Office.Interop.Excel;
using static Midoliy.Office.Interop.FontStyle;
using System.Diagnostics;

namespace TestConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            //using (var app = Excel.BlankWorkbook())
            //{
            //    app.Visibility = AppVisibility.Visible;
            //
            //    var book = app[1];
            //    var sheet = book[1];
            //    sheet["A1"].Value = 100;
            //    sheet["B1"].Value = "Test String";
            //    sheet["B2"].Value = "Test String2";
            //    app[1][1][1, 1].Value = 100;
            //    app[1][1]["C1"].Paste(app[1][1][1, 1]);
            //
            //    // ============================================================================
            //    //     ↓↓↓↓↓    ver 0.0.5.4 追加分    ↓↓↓↓↓
            //    //
            //    var a1 = app.Workbooks(1).Worksheets(1).Cells("A1");
            //
            //    // フォントサイズの変更
            //    a1.Font.Size = 24;
            //
            //    // フォントスタイル変更
            //    a1.Font.Style = Bold | Italic | Shadow | Strikethrough | Subscript | DoubleUnderline;
            //
            //    // セルの削除機能
            //    a1.Delete(DeleteShiftDirection.Up);
            //}

            //using (var app = Excel.BlankWorkbook())
            //{
            //    app.Visibility = AppVisibility.Visible;
            //
            //    var book = app[1];
            //    var sheet = book[1];
            //
            //    // ============================================================================
            //    //     ↓↓↓↓↓    ver 0.0.5.5 追加分    ↓↓↓↓↓
            //    //
            //    var a1k1 = app.Workbooks(1).Worksheets(1).Ranges("A1:K1");
            //
            //    a1k1.Font.Size = 24;
            //    a1k1.Font.Style = Bold | Italic | Shadow | Strikethrough | Subscript | DoubleUnderline;
            //    a1k1.Value = 100;
            //
            //    // A1:K1 に Range を挿入
            //    a1k1.Insert(direction: InsertShiftDirection.Down, origin: InsertFormatOrigin.FromRightOrBelow);
            //
            //    app.Workbooks(1).Worksheets(1).Cells("A1").Value = 200;
            //}

            //using (var app = Excel.BlankWorkbook())
            //{
            //    app.Visibility = AppVisibility.Visible;
            //
            //    var book = app[1];
            //    var sheet = book[1];
            //
            //    // ============================================================================
            //    //     ↓↓↓↓↓    ver 0.0.5.6 追加分    ↓↓↓↓↓
            //    //
            //    var a1k2 = sheet.Ranges("A1:K2");
            //    a1k2.Value = 100;
            //
            //    // 行番号の取得
            //    sheet["A3"].Value = a1k2.Row;
            //    // A1:K2 の範囲に a を代入
            //    a1k2.Rows.Value = "a";
            //    // 行を非表示
            //    //a1k2.Rows.Hidden = true;
            //
            //    // 列番号の取得
            //    sheet["A4"].Value = a1k2.Column;
            //    // A1:K2 の範囲に b を代入
            //    a1k2.Columns.Value = "b";
            //    // 列を非表示
            //    //a1k2.Columns.Hidden = true;
            //
            //    // A列の下端セルを取得
            //    var down = sheet["A1"].End();
            //    // 1行目の右端セルを取得
            //    var right = sheet["A1"].End(Direction.Right);
            //}

        }
    }

    internal static class Native
    {
        [DllImport("ole32.dll")]
        public static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable prot);
    }
}
