using System;
using System.Collections.Generic;

namespace Midoliy.Office.Interop
{
    public interface IExcelRow : IExcelRange
    {
        bool Hidden { get; set; }
        new IEnumerator<IExcelRange> GetEnumerator();
    }
    public interface IExcelRows : IExcelRange
    {
        bool Hidden { get; set; }
        new IEnumerator<IExcelRow> GetEnumerator();
    }
}
