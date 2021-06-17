using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using MsExcel = Microsoft.Office.Interop.Excel;

namespace Midoliy.Office.Interop.Objects
{
    internal struct ExcelRange : IExcelRange, IExcelRow, IExcelRows, IExcelColumns, IExcelColumn
    {
        public int Height => (int)_range.Height;
        public int Width => (int)_range.Width;
        public int RowHeight => (int)_range.RowHeight;
        public int ColumnWidth => (int)_range.ColumnWidth;
        public dynamic Value
        {
            get => _range.Value;
            set => _range.Value = value;
        }

        public dynamic Formula
        {
            get => _range.Formula;
            set => _range.Formula = value;
        }

        public bool Hidden
        {
            get => (bool)_range.Hidden;
            set => _range.Hidden = value;
        }

        public string Address => _range.Address;
        public int Row => _range.Row;
        public IExcelRows Rows => new ExcelRange(_range.Rows, _registerAutoDispose);
        public int Column => _range.Column;
        public IExcelColumns Columns => new ExcelRange(_range.Columns, _registerAutoDispose);
        public IRangeFont Font => new RangeFont(_range.Font);
        public IInterior Interior => new Interior(_range.Interior);
        public IBorders Borders => new Borders(_range.Borders);

        public void Activate() => _range.Activate();
        public void Select() => _range.Select();
        public bool Copy() => (bool)_range.Copy();

        public bool Paste(IExcelRange from, PasteType type, PasteOperation operation, bool skipBlanks, bool transpose)
        {
            if (!from.Copy())
                return false;
            
            return (bool)_range.PasteSpecial(
                Paste: (MsExcel.XlPasteType)type,
                Operation: (MsExcel.XlPasteSpecialOperation)operation,
                SkipBlanks: skipBlanks,
                Transpose: transpose);
        }

        public bool Insert(InsertShiftDirection direction = InsertShiftDirection.Down, InsertFormatOrigin origin = InsertFormatOrigin.FromRightOrBelow)
            => (bool)_range.Insert(direction, origin);

        public bool Delete(DeleteShiftDirection direction)
            => (bool)_range.Delete((MsExcel.XlDeleteShiftDirection)direction);

        public void Clear() => _range.Clear();

        public IExcelRange End(Direction direction = Direction.Down) => new ExcelRange(_range.End[(MsExcel.XlDirection)direction], _registerAutoDispose);

        internal ExcelRange(MsExcel.Range range, Action<IExcelRange> registerAutoDispose)
        {
            _range = range;
            _disposedValue = false;
            _registerAutoDispose = registerAutoDispose;
        }

        private MsExcel.Range _range;
        private Action<IExcelRange> _registerAutoDispose;

        #region IDisposable Support
        private bool _disposedValue;

        private void Dispose(bool disposing)
        {
            if (_disposedValue)
                return;

            if (disposing && _range != null)
            {
                try { while (0 < Marshal.ReleaseComObject(_range)) { } } catch { }
                _range = null;
                GC.Collect();
            }

            _disposedValue = true;
        }

        public void Dispose()
            => Dispose(true);

        #endregion

        IEnumerator<IExcelRange> IExcelRange.GetEnumerator()
        {
            var autoDispose = _registerAutoDispose;
            return _range
                .Cast<MsExcel.Range>()
                .Select(r => new ExcelRange(r, autoDispose) as IExcelRange)
                .GetEnumerator();
        }

        IEnumerator<IExcelRow> IExcelRows.GetEnumerator()
        {
            var autoDispose = _registerAutoDispose;
            return _range
                .Rows
                .Cast<MsExcel.Range>()
                .Select(r => new ExcelRange(r, autoDispose) as IExcelRow)
                .GetEnumerator();
        }

        IEnumerator<IExcelRange> IExcelRow.GetEnumerator()
        {
            var autoDispose = _registerAutoDispose;
            return _range
                .Columns
                .Cast<MsExcel.Range>()
                .Select(r => new ExcelRange(r, autoDispose) as IExcelRange)
                .GetEnumerator();
        }

        IEnumerator<IExcelColumn> IExcelColumns.GetEnumerator()
        {
            var autoDispose = _registerAutoDispose;
            return _range
                .Columns
                .Cast<MsExcel.Range>()
                .Select(r => new ExcelRange(r, autoDispose) as IExcelColumn)
                .GetEnumerator();
        }

        IEnumerator<IExcelRange> IExcelColumn.GetEnumerator()
        {
            var autoDispose = _registerAutoDispose;
            return _range
                .Rows
                .Cast<MsExcel.Range>()
                .Select(r => new ExcelRange(r, autoDispose) as IExcelRange)
                .GetEnumerator();
        }
    }
}
