using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Midoliy.Office.Interop
{
    public interface IExcelRange : IDisposable
    {
        dynamic Value { get; set; }
        dynamic Formula { get; set; }

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

        void Clear();
    }
}
