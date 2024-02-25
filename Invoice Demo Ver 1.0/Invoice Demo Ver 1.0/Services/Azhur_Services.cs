using Invoice_Demo_Ver_1._0.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Invoice_Demo_Ver_1._0.Services
{
    public static class Azhur_Services
    {
        public static List<Azhur> Azhur_Data = new List<Azhur>();
        public static void GetTableData(ExcelWorksheet worksheet)
        {
            for (int row = 3; row <= worksheet.Dimension.End.Row; row++)
            {
                if (worksheet.Cells[row, 17].Value == null)
                    break;
                HandleObjectData(worksheet, row);
            }
        }

        public static void HandleObjectData(ExcelWorksheet worksheet, int row)
        {
#pragma warning disable CS8604
#pragma warning disable CS8600
            string Id = worksheet.Cells[row, 17].Value.ToString();
            string DocumentType = worksheet.Cells[row, 18].Value.ToString();

            long DocumentNum = long.Parse(worksheet.Cells[row, 19].Value.ToString());

            DateTime date = DateTime.Parse(worksheet.Cells[row, 20].Value.ToString());
            DateOnly DocumentDate = DateOnly.FromDateTime(date);

            string Name = worksheet.Cells[row, 21].Value.ToString();

            Decimal NoTax = Decimal.Parse(worksheet.Cells[row, 22].Value.ToString());
            Decimal TaxBase = Decimal.Parse(worksheet.Cells[row, 23].Value.ToString());
            Decimal VAT_Base = Decimal.Parse(worksheet.Cells[row, 24].Value.ToString());
            
            string Article = worksheet.Cells[row, 25].Value.ToString();

            Azhur_Data.Add(new Azhur(Id, DocumentType, DocumentNum, DocumentDate, Name, NoTax, TaxBase, VAT_Base, Article));
#pragma warning restore CS8604
#pragma warning restore CS8600
        }

        public static void WriteAzhurTable(ExcelWorksheet worksheet)
        {
            int row = 2;

            foreach (var AzhurObject in Azhur_Data)
            {
                PrintObjectData(worksheet, AzhurObject, row, 17);
                row++;
            }
        }

        public static void Missing(ExcelWorksheet worksheet)
        {
            var missing = Azhur_Data.Where(a => !NRA_Services.NRA_Data.Any(n => n.DocumentNum == a.DocumentNum)).ToList();

            int row = 2;
            foreach (var document in missing)
            {
                PrintObjectData(worksheet, document, row, 1);
                row++;
            }
        }

        public static void PrintObjectData(ExcelWorksheet worksheet, Azhur document, int row, int col)
        {
            worksheet.Cells[row, col++].Value = document.Id;
            worksheet.Cells[row, col++].Value = document.DocumentType;
            worksheet.Cells[row, col++].Value = document.DocumentNum;
            worksheet.Cells[row, col++].Value = document.DocumentDate;
            worksheet.Cells[row, col++].Value = document.Name;
            worksheet.Cells[row, col++].Value = document.NoTax;
            worksheet.Cells[row, col++].Value = document.TaxBase;
            worksheet.Cells[row, col++].Value = document.VAT_Base;
            worksheet.Cells[row, col++].Value = document.Article;

            row++;
        }

        public static void PrintHeader(ExcelWorksheet worksheet, int col)
        {
            worksheet.Cells[1, col++].Value = "ИН по ДДС на контрагента";
            worksheet.Cells[1, col++].Value = "Вид на документа";
            worksheet.Cells[1, col++].Value = "Номер на документа";
            worksheet.Cells[1, col++].Value = "Дата на документа";
            worksheet.Cells[1, col++].Value = "Име на контрагента";
            worksheet.Cells[1, col++].Value = "БЕЗ";
            worksheet.Cells[1, col++].Value = "ДО";
            worksheet.Cells[1, col++].Value = "ДДС";
            worksheet.Cells[1, col++].Value = "Счетов. статия от-до/месец";
        }
    }
}
