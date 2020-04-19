using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Midoliy.Office.Interop
{
    public interface IWorksheet : IDisposable
    {
        SheetVisiblity Visiblity { get; set; }
        string Name { get; }

        IExcelRange this[int row, int col] { get; }
        IExcelRange this[string address] { get; }
        IExcelRange this[string begin, string end] { get; }

        IExcelRange Cells(int row, int col);
        IExcelRange Cells(string address);
        IExcelRange Ranges(string range);
        IExcelRange Ranges(string begin, string end);

        void Hide();
        void Unhide();
        void Rename(string name);
        void Delete();
    }
}
