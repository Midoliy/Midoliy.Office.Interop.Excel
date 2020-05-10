using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MsExcel = Microsoft.Office.Interop.Excel;

namespace Midoliy.Office.Interop.Objects
{
    public class RangeFont : IRangeFont
    {
        public double Size { get => (double)_font.Size; set => _font.Size = value; }

        public FontStyle FontStyle 
        {
            get
            {
                var style = FontStyle.None;

                if ((bool)_font.Bold)
                    style |= FontStyle.Bold;

                if ((bool)_font.Italic)
                    style |= FontStyle.Italic;

                if ((bool)_font.Shadow)
                    style |= FontStyle.Shadow;

                if ((bool)_font.Strikethrough)
                    style |= FontStyle.Strikethrough;

                if ((bool)_font.Subscript)
                    style |= FontStyle.Subscript;

                if ((bool)_font.Superscript)
                    style |= FontStyle.Superscript;

                var underline = (MsExcel.XlUnderlineStyle)_font.Underline;
                if(underline == MsExcel.XlUnderlineStyle.xlUnderlineStyleSingle)
                    style |= FontStyle.SingleUnderline;

                if (underline == MsExcel.XlUnderlineStyle.xlUnderlineStyleDouble)
                    style |= FontStyle.DoubleUnderline;

                return style;
            }
            set
            {
                if (value == FontStyle.None)
                {
                    _font.Bold = false;
                    _font.Italic = false;
                    _font.Shadow = false;
                    _font.Strikethrough = false;
                    _font.Subscript = false;
                    _font.Superscript = false;
                    _font.Underline = MsExcel.XlUnderlineStyle.xlUnderlineStyleNone;
                }
                else
                {
                    if (value.HasFlag(FontStyle.Bold))
                        _font.Bold = true;

                    if (value.HasFlag(FontStyle.Italic))
                        _font.Italic = true;

                    if (value.HasFlag(FontStyle.Shadow))
                        _font.Shadow = true;

                    if (value.HasFlag(FontStyle.Strikethrough))
                        _font.Strikethrough = true;

                    if (value.HasFlag(FontStyle.Subscript))
                        _font.Subscript = true;

                    if (value.HasFlag(FontStyle.Superscript))
                        _font.Superscript = true;

                    if (value.HasFlag(FontStyle.SingleUnderline))
                        _font.Underline = MsExcel.XlUnderlineStyle.xlUnderlineStyleSingle;

                    if (value.HasFlag(FontStyle.DoubleUnderline))
                        _font.Underline = MsExcel.XlUnderlineStyle.xlUnderlineStyleDouble;
                }
            }
        }

        public RangeFont(MsExcel.Font font)
            => _font = font;

        private readonly MsExcel.Font _font;
    }
}
