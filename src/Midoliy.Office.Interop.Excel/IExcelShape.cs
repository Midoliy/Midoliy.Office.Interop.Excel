using System;
using System.Collections.Generic;

namespace Midoliy.Office.Interop
{
    public interface IExcelShape
    {
        void Apply();
        void Delete();

        string Name { get; set; }
        string Title { get; set; }
        string AlternativeText { get; set; }
    }
}
