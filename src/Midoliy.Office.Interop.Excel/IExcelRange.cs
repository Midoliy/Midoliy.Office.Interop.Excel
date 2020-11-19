using System;

namespace Midoliy.Office.Interop
{
    using static Midoliy.Office.Interop.DeleteShiftDirection;
    using static Midoliy.Office.Interop.InsertShiftDirection;
    using static Midoliy.Office.Interop.InsertFormatOrigin;
    public interface IExcelRange : IDisposable
    {
        dynamic Value { get; set; }
        dynamic Formula { get; set; }
        IRangeFont Font { get; }

        /// <summary>
        /// クリップボードにコピーする
        /// </summary>
        /// <returns>true: 処理成功</returns>
        bool Copy();

        /// <summary>
        /// 対象のセルをコピー＆ペーストする
        /// </summary>
        /// <param name="from">コピー元セル情報</param>
        /// <param name="type">貼り付け形式</param>
        /// <param name="operation">演算方法</param>
        /// <param name="skipBlanks">空白セルを無視するか</param>
        /// <param name="transpose">行列を入れ替えるか</param>
        /// <returns>true: 処理成功</returns>
        bool Paste(IExcelRange from, PasteType type = PasteType.All, PasteOperation operation = PasteOperation.None, bool skipBlanks = false, bool transpose = false);

        /// <summary>
        /// 対象のセルまたは列範囲にセル/列を挿入する
        /// </summary>
        /// <param name="direction">挿入後、元の範囲を右方向と下方向のどちらに移動するか指定</param>
        /// <param name="origin">書式をコピーしてくる方向を指定</param>
        /// <returns>true: 処理成功</returns>
        bool Insert(InsertShiftDirection direction = Down, InsertFormatOrigin origin = FromRightOrBelow);

        /// <summary>
        /// 対象のセルを削除する
        /// </summary>
        /// <param name="direction">セルを削除したあとのシフト方向</param>
        /// <returns>true: 処理成功</returns>
        bool Delete(DeleteShiftDirection direction = Left);

        void Clear();
    }
}
