namespace Midoliy.Office.Interop
{
    public enum PasteType
    {
        /// <summary>値</summary>
        Values = -4163,
        /// <summary>コメント</summary>
        Comments = -4144,
        /// <summary>数式</summary>
        Formulas = -4123,
        /// <summary>書式</summary>
        Formats = -4122,
        /// <summary>すべて</summary>
        All = -4104,
        /// <summary>入力規則</summary>
        Validation = 6,
        /// <summary>罫線を除くすべて</summary>
        AllExceptBorders = 7,
        /// <summary>列幅</summary>
        ColumnWidths = 8,
        /// <summary>数式と数値の書式</summary>
        FormulasAndNumberFormats = 11,
        /// <summary>値と数値の書式</summary>
        ValuesAndNumberFormats = 12,
        /// <summary>コピー元のテーマを使用してすべて貼り付け</summary>
        AllUsingSourceTheme = 13,
        /// <summary>すべての結合されている条件付き書式</summary>
        AllMergingConditionalFormats = 14
    }
}
