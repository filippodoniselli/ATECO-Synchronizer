using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Linq;

[assembly: log4net.Config.XmlConfigurator(Watch = true)]

namespace AtecoSynchronizer
{
    public class Program
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        static void Main()
        {
            log.Info("Program Running");
            string IRP_path = Directory.GetCurrentDirectory() + ConfigurationManager.AppSettings["IRP"], AIRAP_path = Directory.GetCurrentDirectory() + ConfigurationManager.AppSettings["AIRAP"];
            string Cod_Att_path = Directory.GetCurrentDirectory() + ConfigurationManager.AppSettings["Cod_Att"];
            string[] IRP_files = Directory.GetFiles(IRP_path), AIRAP_files = Directory.GetFiles(AIRAP_path);
            if (File.Exists(Cod_Att_path) == false || IRP_files == null || AIRAP_files == null)
            {
                log.Warn("Files not found in specified directory");
                return;
            }
            if (IRP_files.Length != AIRAP_files.Length)
            {
                log.Warn("Uncorrect number of input files");
                return;
            }
            int i;
            if (WellFormed(Cod_Att_path) == false)
                return;
            List<Tuple<string, string>> Cod_Att = Convert(Cod_Att_path);
            for (i = 0; i < IRP_files.Length; i++)
            {
                if (WellFormed(IRP_files[i]) == true && WellFormed(AIRAP_files[i]) == true)
                {
                    List<Tuple<string, string>> IRP = Convert(IRP_files[i]), AIRAP = Convert(AIRAP_files[i]);
                    DirectoryInfo result = new DirectoryInfo(Directory.GetCurrentDirectory() + "\\_Result_");
                    result.Create();
                }
            }
        }
        public static bool WellFormed(string path)
        {
            FileStream input = new FileStream(path, FileMode.Open, FileAccess.Read);
            if (path.EndsWith(".xls"))
            {
                HSSFWorkbook workbook = new HSSFWorkbook(input);
                ISheet sheet = workbook.GetSheetAt(0);
                if (CheckingInstructionSet(sheet) == false)
                {
                    log.Warn("Error in " + path);
                    return false;
                }
                else
                    return true;
            }
            else if (path.EndsWith(".xlsx"))
            {
                XSSFWorkbook workbook = new XSSFWorkbook(input);
                ISheet sheet = workbook.GetSheetAt(0);
                if (CheckingInstructionSet(sheet) == false)
                {
                    log.Warn("Error in " + path);
                    return false;
                }
                return true;
            }
            else
            {
                log.Warn("Uncorrect extension: " + path);
                return false;
            }
        }
        public static bool CheckingInstructionSet(ISheet sheet)
        {
            int i;
            IRow row = sheet.GetRow(0); ;
            switch (row.PhysicalNumberOfCells)
            {
                case 2:
                    for (i = 0; i <= sheet.LastRowNum; i++)
                    {
                        row = sheet.GetRow(i);
                        if (row.GetCell(0).CellType != CellType.String || row.GetCell(1).CellType != CellType.String || String.IsNullOrWhiteSpace(row.GetCell(0).StringCellValue) == true || String.IsNullOrWhiteSpace(row.GetCell(1).StringCellValue) == true)
                            return false;
                    }
                    break;
                case 4:
                case 6:
                case 8:
                    for (i = 0; i <= sheet.LastRowNum; i++)
                    {
                        row = sheet.GetRow(i);
                        if (row.GetCell(0).StringCellValue == "XX")
                            return true;
                        if (row.GetCell(0).CellType != CellType.String || row.GetCell(2).CellType != CellType.String || String.IsNullOrWhiteSpace(row.GetCell(0).StringCellValue) == true  || String.IsNullOrWhiteSpace(row.GetCell(2).StringCellValue) == true)
                            return false;
                    }
                    break;
                default:
                    return false;
            }
            return true;
        }
        public static List<Tuple<string, string>> Convert(string path)
        {
            FileStream input = new FileStream(path, FileMode.Open, FileAccess.Read);
            if (path.EndsWith(".xls"))
            {
                HSSFWorkbook workbook = new HSSFWorkbook(input);
                ISheet sheet = workbook.GetSheetAt(0);
                return ConversionInstructionSet(sheet);
            }
            else
            {
                XSSFWorkbook workbook = new XSSFWorkbook(input);
                ISheet sheet = workbook.GetSheetAt(0);
                return ConversionInstructionSet(sheet);
            }
        }
        public static List<Tuple<string, string>> ConversionInstructionSet(ISheet sheet)
        {
            var result = new List<Tuple<string, string>>(sheet.LastRowNum);
            int i;
            IRow row = sheet.GetRow(0);
            if (row.PhysicalNumberOfCells == 2)
            {
                for(i=0; i <= sheet.LastRowNum; i++)
                {
                    row = sheet.GetRow(i);
                    var item = new Tuple<string, string>(row.GetCell(0).StringCellValue, row.GetCell(1).StringCellValue);
                    result.Add(item);
                }
            }
            else
            {
                for(i=0; i <= sheet.LastRowNum; i++)
                {
                    row = sheet.GetRow(i);
                    if (row.GetCell(0).StringCellValue == "XX")
                        break;
                    var item = new Tuple<string, string>(row.GetCell(0).StringCellValue, row.GetCell(2).StringCellValue);
                    result.Add(item);
                }
            }
            return result;
        }
    }
}