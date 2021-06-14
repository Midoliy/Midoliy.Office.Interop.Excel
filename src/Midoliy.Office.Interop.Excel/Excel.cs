using Midoliy.Office.Interop.Objects;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using MsExcel = Microsoft.Office.Interop.Excel;

namespace Midoliy.Office.Interop
{
    internal static class Native
    {
        [DllImport("ole32.dll")]
        public static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable prot);
    }
    public static class Excel
    {
        public static Process[] EnumerateProcess()
            => Process.GetProcessesByName("Excel");


        public static IExcelApplication BlankWorkbook()
        {
            var excel = new ExcelApplication();
            _ = excel.BlankWorkbook();
            return excel;
        }

        public static IExcelApplication CreateFrom(string templatePath)
        {
            var excel = new ExcelApplication();
            _ = excel.CreateFrom(templatePath);
            return excel;
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Interoperability", "CA1416:プラットフォームの互換性を検証", Justification = "<保留中>")]
        public static IExcelApplication Attach(int hwnd)
        {
            IRunningObjectTable table = null;
            IEnumMoniker monikers = null;

            try
            {
                if (Native.GetRunningObjectTable(0, out table) != 0 || table == null)
                    throw new Exception("Running object table is not found.");

                table.EnumRunning(out monikers);
                monikers.Reset();

                var container = new IMoniker[1];
                var fetchedMonikers = IntPtr.Zero;
                while (monikers.Next(1, container, fetchedMonikers) == 0)
                {
                    table.GetObject(container[0], out object com);
                    if (com is MsExcel.Workbook)
                    {
                        var wb = (MsExcel.Workbook)com;
                        if (hwnd == wb.Application.Hwnd)
                            return new ExcelApplication(wb.Application, Calculation.Auto);
                    }

                    if (com != null)
                        while (0 < Marshal.ReleaseComObject(com)) { }
                }
            }
            finally
            {
                if (table != null)
                    while (0 < Marshal.ReleaseComObject(table)) { }
                if (monikers != null)
                    while (0 < Marshal.ReleaseComObject(monikers)) { }
            }

            throw new Exception("The HWND is not found.");
        }

        public static IExcelApplication Open(string filePath)
        {
            var excel = new ExcelApplication();
            _ = excel.Open(filePath);
            return excel;
        }
    }
}
