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
            app[1].Worksheets(1)["A1"].Value = 100;
        }
    }
}
```

---

## 使い方

### 1. Workbook を作成する

1. `BlankWorkbook()` : 空の Workbook を作成する

   ```cs
   // 空のWorkbookを作成する
   using (var app = Excel.BlankWorkbook())
   {
       // Workbooks(int) メソッドで Workbook を取得
       //  -> index は 1 始まりなので注意(VBA準拠)
       var workbook = app.Workbooks(index: 1);

       // インデクサでアクセスすることも可能
       var workbook2 = app[1];
   ```


        // do something...

    }
    ```

2. `CreateFrom(string)` : 指定した Excel ファイルを元に新規 Workbook を作成する

   ```cs
   using (var app = Excel.CreateFrom("TemplateExcelFile.xlsx"))
   {
       // Workbooks(int) メソッドで Workbook を取得
       //  -> index は 1 始まりなので注意(VBA準拠)
       var workbook = app.Workbooks(index: 1);

       // インデクサでアクセスすることも可能
       var workbook2 = app[1];

       // do something...

   }
   ```

3. `Open(string)` : 指定した Excel ファイルを開く

   ```cs
   using (var app = Excel.Open("Foo.xlsx"))
   {
       // Workbooks(int) メソッドで Workbook を取得
       //  -> index は 1 始まりなので注意(VBA準拠)
       var workbook = app.Workbooks(index: 1);

       // インデクサでアクセスすることも可能
       var workbook2 = app[1];

       // do something...

   }
   ```

4. `Save()` : `Open(string)` したファイルを上書き保存する

   ```cs
   using (var app = Excel.Open("Foo.xlsx"))
   {
       var workbook = app.Workbooks(index: 1);

       // do something...

       // Save() を呼び出さない限り, 処理内容は保存されないので注意
       workbook.Save();
   }
   ```

   以下のように, 新規で Workbook を作成した場合は, 特に何も起きない

   ```cs
   using (var app = Excel.BlankWorkbook())
   // or using (var app = Excel.CreateFrom("TemplateExcelFile.xlsx"))
   {
       var workbook = app.Workbooks(index: 1);

       // do something...

       // Save() を呼び出しても保存されない
       workbook.Save();
   }
   ```

5. `SaveAs(string)` : Workbook を名前をつけて保存する

   ```cs
   using (var app = Excel.BlankWorkbook())
   {
       var workbook = app.Workbooks(index: 1);

       // do something...

       // SaveAs(string) で Workbook を名前を付けて保存
       workbook.SaveAs("NewExcelWorkbook.xlsx");
   }
   ```

---

### 2. Workbook を操作する

1. `Worksheets(int)` : index の位置にある Worksheet を取得する

   ```cs
   using (var app = Excel.BlankWorkbook())
   {
       var workbook = app.Workbooks(index: 1);

       // Worksheets(int) メソッドで Worksheet を取得
       //  -> index は 1 始まりなので注意(VBA準拠)
       var worksheet = workbook.Worksheets(index: 1);

       // インデクサでアクセスすることも可能
       var worksheet2 = workbook[1];

       // do something...

   }
   ```

2. `Worksheets(string)` : 指定した Worksheet 名に一致する Worksheet を取得する

   ```cs
   using (var app = Excel.BlankWorkbook())
   {
       var workbook = app.Workbooks(index: 1);

       // Worksheets(int) メソッドで Worksheet を取得
       //  -> index は 1 始まりなので注意(VBA準拠)
       var worksheet = workbook.Worksheets(name: "Sheet1");

       // インデクサでアクセスすることも可能
       var worksheet2 = workbook["Sheet1"];

       // do something...

   }
   ```

3. `NewSheet()` : 新しい Worksheet を Workbook に追加する

   ```cs
   using (var app = Excel.BlankWorkbook())
   {
       var workbook = app.Workbooks(index: 1);

       // NewSheet() メソッドで新しい Worksheet を追加
       //  -> 新しく追加した Worksheet は戻り値として得られる
       var newsheet = workbook.NewSheet();

       // 新しい Worksheet は Workbook の最後尾に追加されているので,
       // LINQメソッドの Last() / LastOrDefault() を利用することでも得られる
       var newsheet2 = workbook.Last();

       // do something...

   }
   ```

4. `NewSheet(string)` : 新しい Worksheet をシート名を指定して Workbook に追加する

   ```cs
   using (var app = Excel.BlankWorkbook())
   {
       var workbook = app.Workbooks(index: 1);

       // NewSheet(string) メソッドで新しい Worksheet を追加
       //  -> 新しく追加した Worksheet は戻り値として得られる
       var newsheet = workbook.NewSheet(sheetName: "SampleSheet");

       // 新しい Worksheet は Workbook の最後尾に追加されているので,
       // LINQメソッドの Last() / LastOrDefault() を利用することでも得られる
       var newsheet2 = workbook.Last();

       // インデクサで取得することも可能
       var newsheet3 = workbook["SampleSheet"];

       // do something...

   }
   ```

---

### 3. Worksheet を操作する

1. `Cells(int, int)` :（行インデックス, 列インデックス） で指定した位置にあるセルを取得する

   ```cs
   using (var app = Excel.BlankWorkbook())
   {
       var sheet = app[1][1];

       // Cells(int, int) メソッドでセルにアクセス
       //  -> index は 1 始まりなので注意(VBA準拠)
       var cell = sheet.Cells(1, 1);   // A1

       // インデクサでアクセスすることも可能
       var cell2 = sheet[1, 1];

       // do something...

   }
   ```

1. `Cells(string)` : A1 形式で指定した位置にあるセルを取得する

   ```cs
   using (var app = Excel.BlankWorkbook())
   {
       var sheet = app[1][1];

       // Cells(string) メソッドでセルにアクセス
       var cell = sheet.Cells("A1");

       // インデクサでアクセスすることも可能
       var cell2 = sheet["A1"];

       // do something...

   }
   ```

1. `Ranges(string)` : A1 形式で指定した位置にあるセルを取得する

   ```cs
   using (var app = Excel.BlankWorkbook())
   {
       var sheet = app[1][1];

       // Ranges(string) メソッドでセルにアクセス
       //  -> Cells(string) と使い方はまったく同じ
       var range = sheet.Ranges("A1:A2");

       // インデクサでアクセスすることも可能
       var range2 = sheet["A1:A2"];

       // do something...

   }
   ```

1. `Ranges(string, string)` : A1 形式の（開始アドレス, 終了アドレス）で指定した範囲にあるセルを取得する

   ```cs
   using (var app = Excel.BlankWorkbook())
   {
       var sheet = app[1][1];

       // Ranges(string, string) メソッドでセルに範囲アクセス
       var range = sheet.Ranges("A1", "C3");

       // インデクサでアクセスすることも可能
       var range2 = sheet["A1", "C3"];

       // do something...

   }
   ```

1. `Rows(int)` : 指定した行を行選択する

   ```cs
   using (var app = Excel.BlankWorkbook())
   {
       var sheet = app[1][1];

       // Rows(int) メソッドで行選択する
       //  -> index は 1 始まりなので注意(VBA準拠)
       var row = sheet.Rows(1);  // 1行目

       // do something...

   }
   ```

1. `Rows(int, int)` :（開始行番号, 終了行番号）で指定した行を範囲行選択する

   ```cs
   using (var app = Excel.BlankWorkbook())
   {
       var sheet = app[1][1];

       // Rows(int, int) メソッドで行を範囲行選択する
       //  -> index は 1 始まりなので注意(VBA準拠)
       var row = sheet.Rows(1, 5);  // 1~5行目

       // do something...

   }
   ```

1. `Columns(int)` : 指定した列を列選択する

   ```cs
   using (var app = Excel.BlankWorkbook())
   {
       var sheet = app[1][1];

       // Columns(int) メソッドで列選択する
       //  -> index は 1 始まりなので注意(VBA準拠)
       var column = sheet.Columns(1);  // 1列目 = A列

       // do something...

   }
   ```

1. `Columns(string)` : 指定した列を列選択する

   ```cs
   using (var app = Excel.BlankWorkbook())
   {
       var sheet = app[1][1];

       // Columns(string) メソッドで列選択する
       var column = sheet.Columns("A");

       // do something...

   }
   ```

1. `Columns(int, int)` :（開始列番号, 終了列番号）で指定した列を範囲列選択する

   ```cs
   using (var app = Excel.BlankWorkbook())
   {
       var sheet = app[1][1];

       // Columns(int, int) メソッドで範囲列選択する
       //  -> index は 1 始まりなので注意(VBA準拠)
       var column = sheet.Columns(1, 5);  // 1~5列目 = A~E列

       // do something...

   }
   ```

1. `Columns(string, string)` :（開始列番号, 終了列番号）で指定した列を範囲列選択する

   ```cs
   using (var app = Excel.BlankWorkbook())
   {
       var sheet = app[1][1];

       // Columns(string, string) メソッドで範囲列選択する
       var column = sheet.Columns("A", "E");

       // do something...

   }
   ```

1. `Hide()` : シートを非表示状態にする

   ```cs
   using (var app = Excel.BlankWorkbook())
   {
       // 複数シートがないと非表示にできないため, 新しいシートを作成する
       var newsheet = app[1].NewSheet();

       // Hide() メソッドでシートを非表示状態にする
       newsheet.Hide();

       // do something...

   }
   ```

1. `Unhide()` : シートを表示状態にする

   ```cs
   using (var app = Excel.BlankWorkbook())
   {
       // 複数シートがないと非表示にできないため, 新しいシートを作成する
       var newsheet = app[1].NewSheet();

       // Hide() メソッドでシートを非表示状態にする
       newsheet.Hide();

       // Unhide() メソッドでシートを表示状態にする
       newsheet.Unhide();

       // do something...

   }
   ```

1. `Rename(string)` : 指定したシート名にシート名を変更する

   ```cs
   using (var app = Excel.BlankWorkbook())
   {
       var sheet = app[1][1];

       // Rename(string) メソッドでシート名を変更する
       sheet.Rename("RenamedSheet");

       // do something...

   }
   ```

1. `Delete()` : シートを削除する

   ```cs
   using (var app = Excel.BlankWorkbook())
   {
       // 複数シートがないと削除できないため, 新しいシートを作成する
       var newsheet = app[1].NewSheet();

       // Delete() メソッドでシートを削除する
       newsheet.Delete();

       // do something...

   }
   ```

1. `Nameプロパティ` : シート名を取得する

   ```cs
   using (var app = Excel.BlankWorkbook())
   {
       var sheet = app[1][1];

       // Nameプロパティでシート名を取得
       var sheetName = sheet.Name;

       // do something...

   }
   ```

1. `Visibilityプロパティ` : シートの可視状態を取得／設定する

   ```cs
   using (var app = Excel.BlankWorkbook())
   {
       var sheet = app[1][1];

       // Visibilityプロパティで表示状態を取得
       var visibility = sheet.Visibility;

       // Visibilityプロパティを利用して, 表示状態を変更する
       sheet.Visibility = SheetVisiblity.Visible;      // 表示
       sheet.Visibility = SheetVisiblity.Hidden;       // 非表示
       sheet.Visibility = SheetVisiblity.VeryHidden;   // 完全非表示

       // do something...

   }
   ```

---

### 4. Cell を操作する

Coming soon...