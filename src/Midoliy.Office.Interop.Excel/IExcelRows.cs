using System;

namespace Midoliy.Office.Interop
{
    public interface IExcelRows : IExcelRange
    {
        bool Hidden { get; set; }
    }
}
