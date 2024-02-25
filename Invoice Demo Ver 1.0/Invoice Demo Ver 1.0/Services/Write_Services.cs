using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Invoice_Demo_Ver_1._0.Services
{
    public static class Write_Services
    {
        public static void WriteExcelWorkbook()
        {
            using (var package = new ExcelPackage())
            {
                OriginalWorksheet(package);
                MissingNRAWorksheet(package);
                MissingAzhurWorksheet(package);
                WrongVATWorksheet(package);
                CancelledDocumentsWorksheet(package);

                File.WriteAllBytes(GetOutputFilePath(), package.GetAsByteArray());
            }
        }

        public static void OriginalWorksheet(ExcelPackage package)
        {
            ExcelWorksheet original_Worksheet = package.Workbook.Worksheets.Add("Оригинална таблица");
            NRA_Services.PrintHeader(original_Worksheet, 1);
            NRA_Services.WriteNRATable(original_Worksheet);
            Azhur_Services.PrintHeader(original_Worksheet, 17);
            Azhur_Services.WriteAzhurTable(original_Worksheet);
        }

        public static void MissingNRAWorksheet(ExcelPackage package)
        {
            ExcelWorksheet missingNRA = package.Workbook.Worksheets.Add("Липсващи от НАП");
            NRA_Services.PrintHeader(missingNRA, 1);
            NRA_Services.Missing(missingNRA);
        }

        public static void MissingAzhurWorksheet(ExcelPackage package)
        {
            ExcelWorksheet missingAzhur = package.Workbook.Worksheets.Add("Липсващи от Ажур");
            Azhur_Services.PrintHeader(missingAzhur, 1);
            Azhur_Services.Missing(missingAzhur);
        }

        public static void WrongVATWorksheet(ExcelPackage package)
        {
            ExcelWorksheet wrongVAT = package.Workbook.Worksheets.Add("Грешно ДДС");
            NRA_Services.PrintHeader(wrongVAT, 1);
            Azhur_Services.PrintHeader(wrongVAT, 17);
            NRA_Services.WrongVAT(wrongVAT);
        }

        public static void CancelledDocumentsWorksheet(ExcelPackage package)
        {
            ExcelWorksheet cancelledDocuments = package.Workbook.Worksheets.Add("Сторнирани");
            NRA_Services.PrintHeader(cancelledDocuments, 1);
            Azhur_Services.PrintHeader(cancelledDocuments, 17);
            NRA_Services.CancelledDocuments(cancelledDocuments);
        }

        public static string GetOutputFilePath()
        {
            string outputFilePath = Path.Combine(GetOutputDirectory(), GetOutputFileName());
            return outputFilePath;
        }

        public static string GetOutputDirectory()
        {
            while (true)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Въведете директорията, в която искате да поставите обработения екселски файл:");
                Console.ResetColor();
                string? outputDirectory = Console.ReadLine();

                if (outputDirectory != null)
                {
                    if (Directory.Exists(outputDirectory))
                        return outputDirectory;

                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"'{outputDirectory}' не е валидна директория.");
                    }
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Опитайте отново.");
                }
            }
        }

        public static string GetOutputFileName()
        {
            string fileName = "Invoice Comparative Analysis";
            string currentDate = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            
            return $"{fileName} {currentDate}.xlsx";
        }
    }
}
