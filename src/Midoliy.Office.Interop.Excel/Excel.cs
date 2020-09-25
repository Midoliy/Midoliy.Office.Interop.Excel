using Midoliy.Office.Interop.Objects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Midoliy.Office.Interop
{
    public static class Excel
    {
        public static IExcelApplication BlankWorkbook()
        {
            var excel = new ExcelApplication();
            var book = excel.BlankWorkbook();
            _ = book.NewSheet();
            return excel;
        }

        public static IExcelApplication CreateFrom(string templatePath)
        {
            var excel = new ExcelApplication();
            var book = excel.CreateFrom(templatePath);
            return excel;
        }

        public static IExcelApplication Open(string filePath)
        {
            var excel = new ExcelApplication();
            _ = excel.Open(filePath);
            return excel;
        }
    }
}
