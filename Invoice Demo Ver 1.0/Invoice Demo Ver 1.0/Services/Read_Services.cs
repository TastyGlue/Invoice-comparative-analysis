﻿using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Invoice_Demo_Ver_1._0.Services
{
    public static class Read_Services
    {
        public static string GetInputFileName()
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
            string filePath = GetInputFileName();
            FileInfo file = new(filePath);
            using (ExcelPackage package = new(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                NRA_Services.GetTableData(worksheet);
                Azhur_Services.GetTableData(worksheet);
            }
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

        public static bool UserInputMainLoop()
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Искате ли да продъжлите програмата?");

            while (true)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("(0 + ENTER) за НЕ | (1 + ENTER) за ДА");
                Console.ResetColor();

                string? userInput = Console.ReadLine();
                if (userInput == "1")
                    return false;
                else if (userInput != "0" || userInput == null)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Невалиден отоговор.");
                }
                else
                    return true;
            }
        }
    }
}
