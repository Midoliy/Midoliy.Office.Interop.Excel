using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using MsExcel = Microsoft.Office.Interop.Excel;

namespace Midoliy.Office.Interop.Objects
{
    internal readonly struct Borders : IBorders
    {
        public Borders(MsExcel.Borders borders)
        {
            _borders = borders;
        }

        private readonly MsExcel.Borders _borders;
    }
}
