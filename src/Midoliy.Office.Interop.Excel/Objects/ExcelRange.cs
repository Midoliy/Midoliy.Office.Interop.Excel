﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using MsExcel = Microsoft.Office.Interop.Excel;

namespace Midoliy.Office.Interop.Objects
{
    internal class ExcelRange : IExcelRange
    {
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

        public void Clear()
            => _range.Clear(); 

        internal ExcelRange(MsExcel.Range range)
        {
            _range = range;
            _disposedValue = false;
        }

        private MsExcel.Range _range;

        #region IDisposable Support
        private bool _disposedValue;

        protected virtual void Dispose(bool disposing)
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
    }
}