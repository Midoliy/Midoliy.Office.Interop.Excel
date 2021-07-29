using System;
using System.Collections.Generic;

namespace Midoliy.Office.Interop
{
    using static Midoliy.Office.Interop.DeleteShiftDirection;
    using static Midoliy.Office.Interop.InsertShiftDirection;
    using static Midoliy.Office.Interop.InsertFormatOrigin;

    public interface IExcelRange //: IDisposable
    {
        double Top { get; }
        double Left { get; }
        int Height { get; }
        int Width { get; }
        int RowHeight { get; set; }
        int ColumnWidth { get; set; }
        dynamic Value { get; set; }
        dynamic Formula { get; set; }
        string Address { get; }
        /// <summary>文字の水平位置.</summary>
        HorizontalAlignment HorizontalAlignment { get; set; }
        /// <summary>文字の垂直位置.</summary>
        VerticalAlignment VerticalAlignment { get; set; }
        /// <summary>折り返して全体を表示.</summary>
        bool WrapText { get; set; }
        /// <summary>縮小して全体を表示.</summary>
        bool ShrinkToFit { get; set; }
        /// <summary>文字の方向. -90 ~ 90の間で指定.</summary>
        int Orientation { get; set; }
        /// <summary>セルの表示形式を指定.</summary>
        string Format { get; set; }

        int Row { get; }
        IExcelRows Rows { get; }
        IExcelRows EntireRow { get; }
        int Column { get; }
        IExcelColumns Columns { get; }
        IExcelColumns EntireColumn { get; }
        IRangeFont Font { get; }
        IInterior Interior { get; }
        IBorders Borders { get; }

        void AutoFit();
        void Activate();
        void Select();

        /// <summary>セルの結合.</summary>
        /// <param name="across">true: 行ごとに結合する.</param>
        void Merge(bool across = false);
        /// <summary>セルの結合を解除.</summary>
        void UnMerge();

        /// <summary>
        /// クリップボードにコピーする
        /// </summary>
        /// <returns>true: 処理成功</returns>
        bool Copy();

        /// <summary>
        /// 現在のセルにクリップボードのデータをペーストする
        /// </summary>
        /// <param name="type">貼り付け形式</param>
        /// <param name="operation">演算方法</param>
        /// <param name="skipBlanks">空白セルを無視するか</param>
        /// <param name="transpose">行列を入れ替えるか</param>
        /// <returns>true: 処理成功</returns>
        bool Paste(PasteType type = PasteType.All, PasteOperation operation = PasteOperation.None, bool skipBlanks = false, bool transpose = false);

        /// <summary>
        /// 対象のセルをコピー＆ペーストする
        /// </summary>
        /// <param name="from">コピー元セル情報</param>
        /// <param name="type">貼り付け形式</param>
        /// <param name="operation">演算方法</param>
        /// <param name="skipBlanks">空白セルを無視するか</param>
        /// <param name="transpose">行列を入れ替えるか</param>
        /// <returns>true: 処理成功</returns>
        bool CopyAndPaste(IExcelRange from, PasteType type = PasteType.All, PasteOperation operation = PasteOperation.None, bool skipBlanks = false, bool transpose = false);

        /// <summary>
        /// 対象のセルまたは列範囲にセル/列を挿入する
        /// </summary>
        /// <param name="direction">挿入後、元の範囲を右方向と下方向のどちらに移動するか指定</param>
        /// <param name="origin">書式をコピーしてくる方向を指定</param>
        /// <returns>true: 処理成功</returns>
        bool Insert(InsertShiftDirection direction = InsertShiftDirection.Down, InsertFormatOrigin origin = FromRightOrBelow);

        /// <summary>
        /// 対象のセルを削除する
        /// </summary>
        /// <param name="direction">セルを削除したあとのシフト方向</param>
        /// <returns>true: 処理成功</returns>
        bool Delete(DeleteShiftDirection direction = DeleteShiftDirection.Left);

        /// <summary>
        /// 指定した方向の最端セルを取得する
        /// </summary>
        /// <param name="direction">検索方向</param>
        /// <returns>最端セルインスタンス</returns>
        IExcelRange End(Direction direction = Direction.Down);

        void Clear();

        IEnumerator<IExcelRange> GetEnumerator();
    }
}
