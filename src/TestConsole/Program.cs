using Midoliy.Office.Interop;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using static Midoliy.Office.Interop.FontStyle;

namespace TestConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            var b = "B".ToColumnNumber();
            var b2 = "B".ToColumnNumber();
            Console.WriteLine(b);
            Console.WriteLine(b2);

            //using (var app = Excel.BlankWorkbook())
            //{
            //    app.Visibility = AppVisibility.Visible;

            //    var book = app[1];
            //    var sheet = book[1];
            //    sheet["A1"].Value = 100;
            //    sheet["B1"].Value = "Test String";
            //    sheet["B2"].Value = "Test String2";
            //    app[1][1][1, 1].Value = 100;
            //    app[1][1]["C1"].Paste(app[1][1][1, 1]);

            //    // ============================================================================
            //    //     ↓↓↓↓↓    ver 0.0.5 追加分    ↓↓↓↓↓
            //    //
            //    var a1 = app.Workbooks(1).Worksheets(1).Cells("A1");

            //    // フォントサイズの変更
            //    a1.Font.Size = 24;

            //    // フォントスタイル変更
            //    a1.Font.Style = Bold | Italic | Shadow | Strikethrough | Subscript | DoubleUnderline;

            //    // セルの削除機能
            //    a1.Delete(ShiftDirection.Up);
            //}
        }
    }
}
