using Invoice_Demo_Ver_1._0.Services;
using OfficeOpenXml;
using Invoice_Demo_Ver_1._0.Models;

namespace Invoice_Demo_Ver_1._0
{
    public class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.Unicode;
            Console.InputEncoding = System.Text.Encoding.Unicode;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            try
            {
                ReadExcelWorksheet();
                Write_Services.WriteExcelWorkbook();
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Вашият обработен екселски файл е готов.");
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Нещо се обърка");
            }
        }

        public static string GetExcelFilePath()
        {
            while (true)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Въведете пълния адрес на екселския файл:");
                Console.ResetColor();
                string? filePath = Console.ReadLine();

                if (filePath != null)
                {
                    if (File.Exists(filePath) && Path.GetExtension(filePath) == ".xlsx")
                    {
                        return filePath;
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"'{filePath}' не е валиден екселски файл.");
                    }
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Опитайте отново.");
                } 
            }
        }

        public static void ReadExcelWorksheet()
        {
            string filePath = GetExcelFilePath();
            FileInfo file = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                NRA_Services.GetTableData(worksheet);
                Azhur_Services.GetTableData(worksheet);
            }
        }
    }
}
