# Midoliy.Office.Interop.Excel

## はじめに

このライブラリは `Excel COM` をより使いやすくするためのラッパーライブラリです。  
通常、`Excel COM` を利用する場合は以下のように煩雑なコードを記述する必要があります。

```cs
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

class Program
{
    static void Main(string[] args)
    {
        var app = new Excel.Application();
        try
        {
            var book = app.Workbooks.Add();
            try
            {
                var sheet = book.Sheets[1];
                try
                {
                    var cell = sheet.Range["A1"] as Excel.Range;
                    try
                    {
                        cell.Value = 100;
                    }
                    finally
                    {
                        while (0 < Marshal.ReleaseComObject(cell)) { } 
                    }
                }
                finally
                {
                    while (0 < Marshal.ReleaseComObject(sheet)) { }
                }
            }
            finally
            {
                while (0 < Marshal.ReleaseComObject(book)) { }
            }
        }
        finally
        {
            app.Visible = true;
            while (0 < Marshal.ReleaseComObject(app)) { }
        }
    }
}
```

`Midoliy.Office.Interop.Excel` を利用することで以下のように簡潔に記述することが可能となります。

```cs
using Midoliy.Office.Interop;

class Program
{
    static void Main(string[] args)
    {
        using (var app = Excel.BlankWorkbook())
        {
            // appのDispose後にExcelを表示する
            app.Visibility = AppVisibility.Visible;

            // パターン(1)
            app.Workbooks(1).Worksheets(1).Cells("A1").Value = 100;

            // パターン(2)
            app[1][1]["A1"].Value = 100;

            // パターン(1) と パターン(2)を複合させても良い
            app[1].Worksheets(1)["A1"].Valeue = 100;                
        }
    }
}
```

## 使い方

Coming soon...