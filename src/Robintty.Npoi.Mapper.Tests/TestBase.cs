﻿using System;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Robintty.Npoi.Mapper.Tests
{
    /// <summary>
    /// Base class for test classes.
    /// </summary>
    public abstract class TestBase
    {
        /// <summary>
        /// Gets a workbook with 2 sheets ("sheet1" and "sheet2") and with 2 rows in "sheet2":
        /// </summary>
        /// <param name="dateValue"></param>
        /// <param name="stringValue"></param>
        /// <returns></returns>
        protected static IWorkbook GetSimpleWorkbook(DateTime dateValue, string stringValue)
        {
            var workbook = GetEmptyWorkbook();
            workbook.CreateSheet("sheet1");
            var sheet = workbook.CreateSheet("sheet2");
            
            var header = sheet.CreateRow(0);
            header.CreateCell(0).SetCellValue("DateProperty");
            header.CreateCell(1).SetCellValue("StringProperty");
            
            var row = sheet.CreateRow(1);
            row.CreateCell(0).SetCellValue(dateValue);
            row.CreateCell(1).SetCellValue(stringValue);

            return workbook;
        }

        /// <summary>
        /// Gets a workbook with an empty sheet named "sheet1".
        /// </summary>
        /// <returns></returns>
        protected static IWorkbook GetBlankWorkbook()
        {
            var workbook = GetEmptyWorkbook();
            workbook.CreateSheet("sheet1");
            return workbook;
        }

        /// <summary>
        /// Gets a workbook without any sheets.
        /// </summary>
        /// <returns></returns>
        protected static IWorkbook GetEmptyWorkbook() => new XSSFWorkbook();

        protected static IWorkbook WriteAndReadBack(IWorkbook workbook, string fileName = "TempWrite")
        {
            using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }

            return WorkbookFactory.Create(fileName);
        }
    }
}
