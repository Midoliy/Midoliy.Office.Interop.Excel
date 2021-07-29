using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Midoliy.Office.Interop
{
    public interface IWorksheet : IDisposable
    {
        /// <summary>
        /// シートの表示状態
        /// </summary>
        SheetVisiblity Visibility { get; set; }

        /// <summary>
        /// シート名
        /// </summary>
        string Name { get; set; }

        /// <summary>
        /// セル選択
        /// </summary>
        /// <param name="row">行番号：1始まり</param>
        /// <param name="col">列番号：1始まり</param>
        /// <returns>選択したセルオブジェクト</returns>
        IExcelRange this[int row, int col] { get; }

        /// <summary>
        /// セル選択
        /// </summary>
        /// <param name="address">A1形式文字列</param>
        /// <returns>選択したセルオブジェクト</returns>
        IExcelRange this[string address] { get; }

        /// <summary>
        /// セルの範囲選択
        /// </summary>
        /// <param name="begin">セル範囲の開始セル：A1形式文字列</param>
        /// <param name="end">セル範囲の終了セル：A1形式文字列</param>
        /// <returns>選択したセルオブジェクト</returns>
        IExcelRange this[string begin, string end] { get; }

        /// <summary>
        /// 行選択
        /// </summary>
        /// <param name="row">行番号：1始まり</param>
        /// <returns>選択した行オブジェクト</returns>
        IExcelRow Rows(int row);

        /// <summary>
        /// 行選択（複数行）
        /// </summary>
        /// <param name="begin">開始行番号：1始まり</param>
        /// <param name="end">終了行番号：1始まり</param>
        /// <returns>選択した行オブジェクト</returns>
        IExcelRows Rows(int begin, int end);

        /// <summary>
        /// 列選択
        /// </summary>
        /// <param name="col">列番号：1始まり</param>
        /// <returns>選択した列オブジェクト</returns>
        IExcelColumn Columns(int col);

        /// <summary>
        /// 列選択（複数列）
        /// </summary>
        /// <param name="begin">開始列番号：1始まり</param>
        /// <param name="end">終了列番号：1始まり</param>
        /// <returns>選択した列オブジェクト</returns>
        IExcelColumns Columns(int begin, int end);

        /// <summary>
        /// 列選択
        /// </summary>
        /// <param name="col">列番号：A始まり</param>
        /// <returns>選択した列オブジェクト</returns>
        IExcelColumns Columns(string col);

        /// <summary>
        /// 列選択（複数列）
        /// </summary>
        /// <param name="begin">開始列番号：A始まり</param>
        /// <param name="end">終了列番号：A始まり</param>
        /// <returns>選択した列オブジェクト</returns>
        IExcelColumns Columns(string begin, string end);

        /// <summary>
        /// セル選択
        /// </summary>
        /// <param name="row">行番号：1始まり</param>
        /// <param name="col">列番号：1始まり</param>
        /// <returns>選択したセルオブジェクト</returns>
        IExcelRange Cells(int row, int col);

        /// <summary>
        /// セル選択
        /// </summary>
        /// <param name="address">A1形式文字列</param>
        /// <returns>選択したセルオブジェクト</returns>
        IExcelRange Cells(string address);

        /// <summary>
        /// セルの範囲選択
        /// </summary>
        /// <param name="range">A1形式文字列</param>
        /// <returns>選択したセルオブジェクト</returns>
        IExcelRange Ranges(string range);

        /// <summary>
        /// セルの範囲選択
        /// </summary>
        /// <param name="begin">セル範囲の開始セル：A1形式文字列</param>
        /// <param name="end">セル範囲の終了セル：A1形式文字列</param>
        /// <returns>選択したセルオブジェクト</returns>
        IExcelRange Ranges(string begin, string end);

        /// <summary>
        /// シートをアクティブにする
        /// </summary>
        void Activate();

        /// <summary>
        /// シートを複数アクティブにする
        /// </summary>
        void Select();

        /// <summary>
        /// シートを非表示にする
        /// </summary>
        void Hide();

        /// <summary>
        /// シートを表示する
        /// </summary>
        void Unhide();

        /// <summary>
        /// シート名を変更する
        /// </summary>
        /// <param name="name"></param>
        void Rename(string name);

        /// <summary>
        /// シートを削除する
        /// </summary>
        void Delete();

        void Save();
        void SaveAs(string fullpath);
    }
}
