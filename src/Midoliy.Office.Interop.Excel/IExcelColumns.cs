using System;

namespace Midoliy.Office.Interop
{
    public interface IExcelColumns : IExcelRange
    {
        bool Hidden { get; set; }
    }
}
