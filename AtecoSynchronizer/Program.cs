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
            string IRP_path = Directory.GetCurrentDirectory() + ConfigurationManager.AppSettings["IRP"];
            string AIRAP_path = Directory.GetCurrentDirectory() + ConfigurationManager.AppSettings["AIRAP"];
            string Cod_Att_path = Directory.GetCurrentDirectory() + ConfigurationManager.AppSettings["Cod_Att"];
            string[] IRP_files = Directory.GetFiles(IRP_path), AIRAP_files = Directory.GetFiles(AIRAP_path);
            if (!File.Exists(Cod_Att_path) || 
                IRP_files == null ||
                AIRAP_files == null)
            {
                log.Warn("Files not found in specified directory");
                return;
            }
            if (!WellFormed(Cod_Att_path))
            {
                log.Error("Cod_Att bad formed. Impossible to generate output");
                return;
            }
            List<Tuple<string, string>> Cod_Att = Convert(Cod_Att_path);
            int i;
            for(i=0; i < IRP_files.Length; i++)
            {
                if ( WellFormed(IRP_files[i]) && WellFormed(AIRAP_files[i]))
                {
                    List<Tuple<string, string>> IRP = Convert(IRP_files[i]), AIRAP = Convert(AIRAP_files[i]);
                    DirectoryInfo result = new DirectoryInfo(Directory.GetCurrentDirectory() + @"\_Result_");
                    result.Create();
                }
            }
        }
        public static bool WellFormed(string path)
        {
            log.Info("Opening to check if well formed: " + path);
            if (!path.Contains(".xls"))
            {
                log.Warn("Uncorrect extension of:" + path);
                return false;
            }
            ISheet sheet = path.EndsWith(".xls") ?
                new HSSFWorkbook(new FileStream(path, FileMode.Open, FileAccess.Read)).GetSheetAt(0) : 
                new XSSFWorkbook(new FileStream(path, FileMode.Open, FileAccess.Read)).GetSheetAt(0);
            bool response = CheckingInstructionSet(sheet);
            log.Info(response ? "Well formed" : "Bad formed");
            return response;
        }
        public static bool CheckingInstructionSet(ISheet sheet)
        {
            log.Info("Checking file");
            int i;
            bool response = true;
            switch (sheet.GetRow(0).PhysicalNumberOfCells)
            {
                case 2:
                    for (i = 0; i <= sheet.LastRowNum; i++)
                    {
                        if (sheet.GetRow(i).GetCell(0).CellType != CellType.String ||
                            sheet.GetRow(i).GetCell(1).CellType != CellType.String ||
                            String.IsNullOrWhiteSpace(sheet.GetRow(i).GetCell(0).StringCellValue) ||
                            String.IsNullOrWhiteSpace(sheet.GetRow(i).GetCell(1).StringCellValue))
                        {
                            log.Warn($"Error in row {i}");
                            response = false;
                        }
                    }
                    break;
                case 4:
                case 6:
                case 8:
                    for (i = 0; i <= sheet.LastRowNum; i++)
                    {
                        if (sheet.GetRow(i).GetCell(0).StringCellValue != "XX" &&
                            (sheet.GetRow(i).GetCell(0).CellType != CellType.String ||
                            sheet.GetRow(i).GetCell(2).CellType != CellType.String || 
                            String.IsNullOrWhiteSpace(sheet.GetRow(i).GetCell(0).StringCellValue) || 
                            String.IsNullOrWhiteSpace(sheet.GetRow(i).GetCell(2).StringCellValue)))
                        {
                            log.Warn($"Error in row {i}");
                            response = false;
                        }
                    }
                    break;
                default:
                    response = false;
                    break;
            }
            return response;
        }
        public static List<Tuple<string, string>> Convert(string path)
        {
            log.Info("Opening to convert " + path);
            return path.EndsWith(".xls") ? 
                ConversionInstructionSet(new HSSFWorkbook(new FileStream(path, FileMode.Open, FileAccess.Read)).GetSheetAt(0)) : 
                ConversionInstructionSet(new XSSFWorkbook(new FileStream(path, FileMode.Open, FileAccess.Read)).GetSheetAt(0));
        }
        public static List<Tuple<string, string>> ConversionInstructionSet(ISheet sheet)
        {
            log.Info("Converting file");
            List<Tuple< string, string>> result = new List<Tuple<string, string>>();
            foreach (IRow row in sheet)
            {
                if (row.PhysicalNumberOfCells == 2)
                    result.Add(new Tuple<string, string> (row.GetCell(0).StringCellValue, row.GetCell(1).StringCellValue));
                else if (row.GetCell(0).StringCellValue != "XX")
                    result.Add(new Tuple<string, string>(row.GetCell(0).StringCellValue, row.GetCell(2).StringCellValue));
            }
            log.Info("Converted");
            return result;
        }
    }
}