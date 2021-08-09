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
        public void Apply() => _shape.Apply();
        public void Delete() => _shape.Delete();

        public string Name { get => _shape.Name; set => _shape.Name = value; }
        public string Title { get => _shape.Title; set => _shape.Title = value; }
        public string AlternativeText { get => _shape.AlternativeText; set => _shape.AlternativeText = value; }

        internal ExcelShape(MsExcel.Shape shape)
        {
            _shape = shape;
        }

        internal readonly MsExcel.Shape _shape;
    }
}
