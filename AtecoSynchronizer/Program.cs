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
            log.Info("PROGRAM RUNNING");
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
            List<KeyValuePair<string, string>> Cod_Att = Converter(Cod_Att_path);
            DirectoryInfo result = new DirectoryInfo(Directory.GetCurrentDirectory() + @"\_Result_");
            result.Create();
            foreach (string IRP in IRP_files)
            {
                string AIRAP = Array.Find(AIRAP_files, str => (Path.GetFileName(IRP).Replace("IRP", "").ToLower()).Contains(Path.GetFileNameWithoutExtension(str).Replace("AIRAP", "").ToLower()));
                if (AIRAP == null)
                    log.Warn($"Not found AIRAP for {Path.GetFileName(IRP)}");
                if(AIRAP != null && WellFormed(IRP) && WellFormed(AIRAP))
                {
                    log.Info($"Trying to correlate {Path.GetFileName(IRP)} e {Path.GetFileName(AIRAP)}");
                    List<KeyValuePair<string, string>> IRPconv = Converter(IRP), AIRAPconv = Converter(AIRAP);
                    List<Tuple<string, string, string, string>> Final = Correlate(IRPconv, AIRAPconv, Cod_Att);
                    if (Final.Count != 0)
                    {
                        log.Info("Correlation completed. Creating resulting file");
                        OutputFile(Final).Write(new FileStream(result.ToString() + @"\" + Path.GetFileNameWithoutExtension(IRP) + ".xlsx", FileMode.OpenOrCreate, FileAccess.Write));
                        log.Info($"Resulting file of {Path.GetFileNameWithoutExtension(IRP)} created!\n");
                    }
                }
            }
        }
        public static bool WellFormed(string path)
        {
            log.Info($"Opening to check if well formed: {Path.GetFileName(path)}");
            if (!path.Contains(".xls"))
            {
                log.Warn($"Uncorrect extension of: {Path.GetFileName(path)}");
                return false;
            }
            ISheet sheet = path.EndsWith(".xls") ?
                new HSSFWorkbook(new FileStream(path, FileMode.Open, FileAccess.Read)).GetSheetAt(0) : 
                new XSSFWorkbook(new FileStream(path, FileMode.Open, FileAccess.Read)).GetSheetAt(0);
            bool response = CheckingInstructionSet(sheet);
            log.Info(response ? "WELL FORMED" : "BAD FORMED");
            return response;
        }
        public static bool CheckingInstructionSet(ISheet sheet)
        {
            log.Info("Checking file");
            bool response = true;
            switch (sheet.GetRow(0).PhysicalNumberOfCells)
            {
                case 2:
                    foreach (IRow row in sheet)
                    {
                        if (row.GetCell(0).CellType != CellType.String ||
                            row.GetCell(1).CellType != CellType.String ||
                            String.IsNullOrWhiteSpace(row.GetCell(0).StringCellValue) ||
                            String.IsNullOrWhiteSpace(row.GetCell(1).StringCellValue))
                        {
                            log.Warn($"Error in row {row.RowNum}");
                            response = false;
                        }
                    }
                    break;
                case 4:
                case 6:
                case 8:
                    foreach (IRow row in sheet)
                    {
                        if (row.GetCell(0).StringCellValue != "XX" &&
                            (row.GetCell(0).CellType != CellType.String ||
                            row.GetCell(2).CellType != CellType.String || 
                            String.IsNullOrWhiteSpace(row.GetCell(0).StringCellValue) || 
                            String.IsNullOrWhiteSpace(row.GetCell(2).StringCellValue)))
                        {
                            log.Warn($"Error in row {row.RowNum}");
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
        public static List<KeyValuePair<string, string>> Converter(string path)
        {
            log.Info($"Opening to convert {Path.GetFileName(path)}");
            return path.EndsWith(".xls") ? 
                ConversionInstructionSet(new HSSFWorkbook(new FileStream(path, FileMode.Open, FileAccess.Read)).GetSheetAt(0)) : 
                ConversionInstructionSet(new XSSFWorkbook(new FileStream(path, FileMode.Open, FileAccess.Read)).GetSheetAt(0));
        }
        public static List<KeyValuePair<string, string>> ConversionInstructionSet(ISheet sheet)
        {
            log.Info("Converting file");
            List<KeyValuePair<string, string>> result = new List<KeyValuePair<string, string>>();
            foreach (IRow row in sheet)
            {
                if (row.PhysicalNumberOfCells == 2)
                    result.Add(new KeyValuePair<string, string> (row.GetCell(0).StringCellValue, row.GetCell(1).StringCellValue));
                else if (row.GetCell(0).StringCellValue != "XX")
                    result.Add(new KeyValuePair<string, string>(row.GetCell(0).StringCellValue, row.GetCell(2).StringCellValue));
            }
            log.Info("CONVERTED");
            return result;
        }
        public static List<Tuple<string, string, string, string>> Correlate(List<KeyValuePair<string, string>> irp, List<KeyValuePair<string, string>> airap, List<KeyValuePair<string, string>> cod_att)
        { 
            var result = new List<Tuple<string, string, string, string>>();
            foreach (KeyValuePair<string, string> couple in irp)
            {
                if (airap.Find(value => value.Key == couple.Value).Value != null && cod_att.Find(value => value.Key == couple.Key).Value != null)
                {
                    result.Add(new Tuple<string, string, string, string>(couple.Key,
                                                                        cod_att.Find(value => value.Key == couple.Key).Value,
                                                                        couple.Value,
                                                                        airap.Find(value => value.Key == couple.Value).Value));
                }
                else
                {
                    log.Warn($"IMPOSSIBLE to correlate elements in IRP row number {irp.IndexOf(couple) + 1}\n");
                    result.Clear();
                    return result;
                }
            }
            return result;
        }
        public static IWorkbook OutputFile(List<Tuple<string, string, string, string>> result)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();
            foreach (Tuple<string, string, string, string> elements in result)
            {
                IRow row = sheet.CreateRow(result.IndexOf(elements));
                ICell cell0 = row.CreateCell(0), cell1 = row.CreateCell(1), cell2 = row.CreateCell(2), cell3 = row.CreateCell(3);
                cell0.SetCellValue(elements.Item1);
                cell1.SetCellValue(elements.Item2);
                cell2.SetCellValue(elements.Item3);
                if (result.IndexOf(elements) != 0)
                    cell3.SetCellValue(Convert.ToDouble(elements.Item4));
                else
                    cell3.SetCellValue(elements.Item4);
            }
            sheet.AutoSizeColumn(1);
            return workbook;
        }
    }
}