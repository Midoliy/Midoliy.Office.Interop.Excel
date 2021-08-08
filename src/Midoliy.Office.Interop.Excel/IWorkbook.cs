using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Midoliy.Office.Interop
{
    public interface IWorkbook : IDisposable, IEnumerable<IWorksheet>
    {
        string Name { get; }
        IWorksheet this[int index] { get; }
        IWorksheet this[string name] { get; }

        IWorksheet Worksheets(int index);
        IWorksheet Worksheets(string name);

        IWorksheet NewSheet();
        IWorksheet NewSheet(string sheetName);

        void Activate();

        void Save();
        void SaveAs(string fullpath);
        void Close(bool saveChanges = false);
    }
}
