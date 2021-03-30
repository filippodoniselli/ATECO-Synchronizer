using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using AtecoSynchronizer;

namespace AtecoSynchronizer_Test
{
    [TestClass]
    public class Program_Test
    {
        [TestMethod]
        public void CheckingIS_2_Rows_True()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();
            IRow row0 = sheet.CreateRow(0), row1 = sheet.CreateRow(1);
            ICell cell00 = row0.CreateCell(0), cell01 = row0.CreateCell(1), cell10 = row1.CreateCell(0), cell11 = row1.CreateCell(1);
            cell00.SetCellValue("-");
            cell01.SetCellValue("-");
            cell10.SetCellValue("1");
            cell11.SetCellValue("1");
            Assert.IsTrue(Program.CheckingInstructionSet(sheet));
        }
        [TestMethod]
        public void CheckingIS_2_Rows_CellTypeError()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();
            IRow row0 = sheet.CreateRow(0), row1 = sheet.CreateRow(1);
            ICell cell00 = row0.CreateCell(0), cell01 = row0.CreateCell(1), cell10 = row1.CreateCell(0), cell11 = row1.CreateCell(1);
            cell00.SetCellValue("-");
            cell01.SetCellValue("-");
            cell10.SetCellValue("1");
            Assert.IsFalse(Program.CheckingInstructionSet(sheet));
        }
        [TestMethod]
        public void CheckingIS_2_Rows_NullOrWhiteError()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();
            IRow row0 = sheet.CreateRow(0), row1 = sheet.CreateRow(1);
            ICell cell00 = row0.CreateCell(0), cell01 = row0.CreateCell(1), cell10 = row1.CreateCell(0), cell11 = row1.CreateCell(1);
            cell00.SetCellValue("-");
            cell01.SetCellValue("-");
            cell10.SetCellValue("1");
            cell11.SetCellValue(" ");
            Assert.IsFalse(Program.CheckingInstructionSet(sheet));
        }
        [TestMethod]
        public void CheckingIS_4_6_8_Rows_True()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();
            IRow row0 = sheet.CreateRow(0), row1 = sheet.CreateRow(1);
            ICell cell00 = row0.CreateCell(0), cell01 = row0.CreateCell(1), cell02 = row0.CreateCell(2), cell03 = row0.CreateCell(3);
            ICell cell10 = row1.CreateCell(0), cell12 = row1.CreateCell(2);
            cell00.SetCellValue("-");
            cell01.SetCellValue("-");
            cell02.SetCellValue("-");
            cell03.SetCellValue("-");
            cell10.SetCellValue("1");
            cell12.SetCellValue("1");
            Assert.IsTrue(Program.CheckingInstructionSet(sheet));
        }
        [TestMethod]
        public void CheckingIS_4_6_8_Rows_CellTypeError()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();
            IRow row0 = sheet.CreateRow(0), row1 = sheet.CreateRow(1);
            ICell cell00 = row0.CreateCell(0), cell01 = row0.CreateCell(1), cell02 = row0.CreateCell(2), cell03 = row0.CreateCell(3);
            ICell cell10 = row1.CreateCell(0), cell12 = row1.CreateCell(2);
            cell00.SetCellValue("-");
            cell01.SetCellValue("-");
            cell02.SetCellValue("-");
            cell03.SetCellValue("-");
            cell10.SetCellValue("1");
            Assert.IsFalse(Program.CheckingInstructionSet(sheet));
        }
        [TestMethod]
        public void CheckingIS_4_6_8_Rows_NullOrWhiteError()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();
            IRow row0 = sheet.CreateRow(0), row1 = sheet.CreateRow(1);
            ICell cell00 = row0.CreateCell(0), cell01 = row0.CreateCell(1), cell02 = row0.CreateCell(2), cell03 = row0.CreateCell(3);
            ICell cell10 = row1.CreateCell(0), cell12 = row1.CreateCell(2);
            cell00.SetCellValue("-");
            cell01.SetCellValue("-");
            cell02.SetCellValue("-");
            cell03.SetCellValue("-");
            cell10.SetCellValue("");
            cell12.SetCellValue("1");
            Assert.IsFalse(Program.CheckingInstructionSet(sheet));
        }
        [TestMethod]
        public void CheckingIS_UncorrectNum_Rows()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();
            IRow row0 = sheet.CreateRow(0);
            ICell cell00 = row0.CreateCell(0), cell01 = row0.CreateCell(1), cell02 = row0.CreateCell(2);
            cell00.SetCellValue("-");
            cell01.SetCellValue("-");
            cell02.SetCellValue("-");
            Assert.IsFalse(Program.CheckingInstructionSet(sheet));
        }
        [TestMethod]
        public void ConversionIS_2_Rows()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();
            IRow row0 = sheet.CreateRow(0), row1 = sheet.CreateRow(1);
            ICell cell00 = row0.CreateCell(0), cell01 = row0.CreateCell(1), cell10 = row1.CreateCell(0), cell11 = row1.CreateCell(1);
            cell00.SetCellValue("-");
            cell01.SetCellValue("-");
            cell10.SetCellValue("1");
            cell11.SetCellValue("1");
            var item1 = new Tuple<string, string>( "-", "-");
            var item2 = new Tuple<string, string>( "1" , "1");
            Assert.IsTrue(Program.ConversionInstructionSet(sheet).IndexOf(item1) == 0 && Program.ConversionInstructionSet(sheet).IndexOf(item2) == 1);
        }
        [TestMethod]
        public void ConversionIS_4_6_8_Rows()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();
            IRow row0 = sheet.CreateRow(0), row1 = sheet.CreateRow(1);
            ICell cell00 = row0.CreateCell(0), cell01 = row0.CreateCell(1), cell02 = row0.CreateCell(2), cell03 = row0.CreateCell(3);
            ICell cell10 = row1.CreateCell(0), cell11 = row1.CreateCell(1), cell12 = row1.CreateCell(2), cell13 = row1.CreateCell(3);
            cell00.SetCellValue("-");
            cell01.SetCellValue("-");
            cell02.SetCellValue("-");
            cell03.SetCellValue("-");
            cell10.SetCellValue("1");
            cell12.SetCellValue("1");
            var item1 = new Tuple<string, string>("-", "-");
            var item2 = new Tuple<string, string>("1", "1");
            Assert.IsTrue(Program.ConversionInstructionSet(sheet).IndexOf(item1) == 0 && Program.ConversionInstructionSet(sheet).IndexOf(item2) == 1);
        }

    }
}