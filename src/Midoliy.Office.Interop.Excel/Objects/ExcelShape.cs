using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Interop.Excel;
using MsExcel = Microsoft.Office.Interop.Excel;

namespace Midoliy.Office.Interop.Objects
{
    internal readonly struct ExcelShape : IExcelShape
    {
        internal ExcelShape(MsExcel.Shape shape)
        {
            _shape = shape;
        }

        private readonly MsExcel.Shape _shape;
    }
}
