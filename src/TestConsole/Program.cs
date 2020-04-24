using Midoliy.Office.Interop;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace TestConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var app = Excel.BlankWorkbook())
            {
                var book = app[1];
                var sheet = book[1];
                sheet["A1"].Value = 100;
                sheet["B1"].Value = "Test String";
                sheet["B2"].Value = "Test String2";
                app[1][1][1, 1].Value = 100;
                app[1][1]["C1"].Paste(app[1][1][1, 1]);
                sheet["B1"].Delete(ShiftDirection.Up);
                app.Visibility = AppVisibility.Visible;
            }
        }
    }
}
