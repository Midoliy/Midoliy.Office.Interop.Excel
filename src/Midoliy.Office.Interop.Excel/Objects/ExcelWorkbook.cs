using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using MsExcel = Microsoft.Office.Interop.Excel;

namespace Midoliy.Office.Interop.Objects
{
    internal class ExcelWorkbook : IWorkbook
    {
        public string Name 
            => _book.Name;

        public IWorksheet this[string name] 
            => _children.First(c => c.Name == name);

        public IWorksheet this[int index] 
            => _children[index - 1];

        public IWorksheet Worksheets(int index)
            => this[index];

        public IWorksheet Worksheets(string name)
            => this[name];

        public IWorksheet NewSheet()
        {
            var sheet = new ExcelWorksheet(_book.Sheets.Add(Count: 1) as MsExcel.Worksheet, onSave: Save, onSaveAs: SaveAs);
            _children.Add(sheet);
            return sheet;
        }

        public IWorksheet NewSheet(string sheetName)
        {
            if (_children.Any(c => c.Name == sheetName))
                throw new AlreadyExistsException(sheetName);
            
            var sheet = NewSheet();
            sheet.Name = sheetName;
            _children.Add(sheet);
            return sheet;
        }

        public void Activate()
        {
            _book.Activate();
            _onActivate?.Invoke(this);
        }

        public void Save()
            => _book.Save();

        public void SaveAs(string fullpath)
            => _book.SaveAs(Path.GetFullPath(fullpath));

        public void Close()
            => _book.Close();

        public IEnumerator<IWorksheet> GetEnumerator()
            => _children.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator()
            => GetEnumerator();

        internal ExcelWorkbook(MsExcel.Workbook book, Action<IWorkbook> onActivate = null)
        {
            _book = book;
            _children = new List<IWorksheet>();
            foreach (MsExcel.Worksheet sheet in _book.Worksheets)
                _children.Add(new ExcelWorksheet(sheet, onSave: Save, onSaveAs: SaveAs));
            _disposedValue = false;
            _onActivate = onActivate;
        }

        private MsExcel.Workbook _book;
        private List<IWorksheet> _children;
        private Action<IWorkbook> _onActivate;

        #region IDisposable Support
        private bool _disposedValue;

        protected virtual void Dispose(bool disposing)
        {
            if (_disposedValue)
                return;

            if (disposing)
            {
                foreach (var sheet in _children)
                    sheet?.Dispose();
                
                try { while (0 < Marshal.ReleaseComObject(_book)) { } } catch { }
                _book = null;

                GC.Collect();
            }

            _disposedValue = true;
        }

        public void Dispose()
            => Dispose(true);
        #endregion
    }
}
