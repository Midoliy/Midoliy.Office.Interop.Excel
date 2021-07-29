using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Midoliy.Office.Interop
{
    public interface IExcelApplication : IDisposable, IEnumerable<IWorkbook>
    {
        AppVisibility Visibility { get; set; }
        Calculation Calculation { get; set; }
        int Count { get; }

        IWorkbook this[int index] { get; }
        IWorkbook this[string name] { get; }

        IWorkbook Workbooks(int index);
        IWorkbook Workbooks(string name);
        IWorkbook Select(int index);
        IWorkbook Select(string name);

        IWorkbook BlankWorkbook();
        IWorkbook CreateFrom(string templatePath);
        IWorkbook Open(string filePath);
    }
}
