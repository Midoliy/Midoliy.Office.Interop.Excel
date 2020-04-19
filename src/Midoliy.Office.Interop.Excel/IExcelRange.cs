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

        void Clear();
    }
}
