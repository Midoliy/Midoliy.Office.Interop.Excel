using System;
using System.Collections.Generic;

namespace Midoliy.Office.Interop
{
    public interface IExcelColumn : IExcelRange
    {
        bool Hidden { get; set; }
        new IEnumerator<IExcelRange> GetEnumerator();
    }
    public interface IExcelColumns : IExcelRange
    {
        bool Hidden { get; set; }
        new IEnumerator<IExcelColumn> GetEnumerator();
    }
}
