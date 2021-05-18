using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using MsExcel = Microsoft.Office.Interop.Excel;

namespace Midoliy.Office.Interop.Objects
{
    internal class ExcelWorksheet : IWorksheet
    {
        public SheetVisiblity Visibility
        { 
            get => (SheetVisiblity)_sheet.Visible; 
            set => _sheet.Visible = (MsExcel.XlSheetVisibility)value;
        }

        public string Name
        {
            get => _sheet.Name;
            set => _sheet.Name = value;
        }

        public IExcelRange this[int row, int col] 
        { 
            get
            {
                if (row < 1 || col < 1)
                    throw new ArgumentOutOfRangeException($"'row' or 'col' is an out-of-range value. ('row' = {row} / 'col' = {col})");

                var range = new ExcelRange(_sheet.Cells[row, col] as MsExcel.Range);
                _trashcan.Add(range);
                return range;
            }
        }

        public IExcelRange this[string address] 
        { 
            get
            {
                var range = new ExcelRange(_sheet.Range[address]);
                _trashcan.Add(range);
                return range;
            }
        }

        public IExcelRange this[string begin, string end]
            => this[$"{begin}:{end}"];

        public IExcelRange Rows(int row)
            => Rows(row, row);

        public IExcelRange Rows(int begin, int end)
            => Ranges($"{begin}:{end}");

        public IExcelRange Columns(int col)
            => Columns(col, col);

        public IExcelRange Columns(int begin, int end)
            => Columns(begin.ToString(), end.ToString());

        public IExcelRange Columns(string col)
            => Columns(col, col);

        public IExcelRange Columns(string begin, string end)
            => Ranges($"{begin}:{end}");

        public IExcelRange Cells(int row, int col)
            => this[row, col];

        public IExcelRange Cells(string address)
            => this[address];

        public IExcelRange Ranges(string range)
            => this[range];

        public IExcelRange Ranges(string begin, string end)
            => this[begin, end];

        public void Hide()
            => _sheet.Visible = MsExcel.XlSheetVisibility.xlSheetHidden;

        public void Unhide()
            => _sheet.Visible = MsExcel.XlSheetVisibility.xlSheetVisible;

        public void Rename(string name)
            => _sheet.Name = name;

        public void Delete()
            => _sheet.Delete();

        internal ExcelWorksheet(MsExcel.Worksheet sheet)
        {
            _sheet = sheet;
            _trashcan = new List<IExcelRange>();
            _disposedValue = false;
        }

        private MsExcel.Worksheet _sheet;
        private List<IExcelRange> _trashcan;

        #region IDisposable Support
        private bool _disposedValue;

        protected virtual void Dispose(bool disposing)
        {
            if (_disposedValue)
                return;

            if (disposing)
            {
                foreach (var range in _trashcan)
                    range?.Dispose();

                if(_sheet != null)
                {
                    try { while (0 < Marshal.ReleaseComObject(_sheet)) { } } catch { }
                    _sheet = null;
                }
                
                GC.Collect();
            }

            _disposedValue = true;
        }

        public void Dispose()
            => Dispose(true);
        #endregion
    }
}
