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
                Read_Services.ReadExcelWorksheet();
                Write_Services.WriteExcelWorkbook();
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Вашият обработен екселски файл е готов.");
            }
            catch (NRAFormatException ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
            }
            catch (AzhurFormatException ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Нещо се обърка.");
                Console.WriteLine($"Грешка: '{ex.Message}'");
            }

            Console.WriteLine("Може да натиснете ENTER за да затворите програмата.");
            Console.ReadLine();
        }
    }
}
