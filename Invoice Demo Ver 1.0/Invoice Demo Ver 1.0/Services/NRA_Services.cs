using Invoice_Demo_Ver_1._0.Models;
using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Invoice_Demo_Ver_1._0.Services
{
    public static class NRA_Services
    {
        public static List<NRA> NRA_Data = new List<NRA>();
        public static List<NRA> anulledNRA = new List<NRA>();
        public static void GetTableData(ExcelWorksheet worksheet)
        {
            for (int row = 13; row <= worksheet.Dimension.End.Row; row++)
            {
                if (worksheet.Cells[row, 1].Value == null)
                    break;
                HandleObjectData(worksheet, row);
            }
        }

        public static void HandleObjectData(ExcelWorksheet worksheet, int row)
        {
#pragma warning disable CS8604
#pragma warning disable CS8600
            try
            {
                string Id = "BG" + worksheet.Cells[row, 1].Value.ToString();
                string Name = worksheet.Cells[row, 2].Value.ToString();
                string Period = worksheet.Cells[row, 3].Value.ToString();

                string DocumentNum = worksheet.Cells[row, 4].Value.ToString();
                DocumentNum = (DocumentNum?.Length < 10) ? DocumentNum.PadLeft(10, '0') : DocumentNum;

                int DocumentType = int.Parse(worksheet.Cells[row, 5].Value.ToString()?.Substring(0, 2));
                if (DocumentType == 9)
                    return;

                DateTime date = DateTime.Parse(worksheet.Cells[row, 6].Value.ToString());
                DateOnly DocumentDate = DateOnly.FromDateTime(date);

                string MerchType = worksheet.Cells[row, 7].Value.ToString();

                Decimal TaxBase = Decimal.Parse(worksheet.Cells[row, 8].Value.ToString());
                Decimal VAT_Base = Decimal.Parse(worksheet.Cells[row, 9].Value.ToString());


                NRA_Data.Add(new NRA(Id, Name, Period, DocumentType, DocumentNum, DocumentDate, MerchType, TaxBase, VAT_Base));
            }
            catch (Exception ex)
            {
                string errorMessageToShow = $"Грешка при форматирането на НАП документ (ред: {row})";
                errorMessageToShow += $"\n{ex.Message}";
                throw new NRAFormatException(errorMessageToShow);
            }
#pragma warning restore CS8604
#pragma warning restore CS8600
        }

        public static void WriteNRATable(ExcelWorksheet worksheet)
        {
            FixCreditSigns();
            //FixInvoiceSigns();

            int row = 2;

            foreach (var NRAobject in NRA_Data)
            {
                PrintObjectData(worksheet, NRAobject, row, 1);
                row++;
            }
        }

        public static void FixCreditSigns()
        {
            foreach (var document in NRA_Data.Where(x => x.DocumentType == 3))
            {
                if (document.TaxBase > 0)
                {
                    document.TaxBase *= -1;
                }

                if (document.VatBase > 0)
                {
                    document.VatBase *= -1;
                }
            }
        }

        public static void FixInvoiceSigns()
        {
            foreach (var document in NRA_Services.NRA_Data.Where(x => x.DocumentType == 1))
            {
                if (document.TaxBase < 0)
                {
                    document.TaxBase *= -1;
                }

                if (document.VatBase < 0)
                {
                    document.VatBase *= -1;
                }
            }
        }

        public static void Missing(ExcelWorksheet worksheet)
        {
            var missing = NRA_Data.Where(n => !Azhur_Services.Azhur_Data.Any(a => (a.DocumentNum == n.DocumentNum) && (a.Id == n.Id))).ToList();

            int row = 2;
            foreach (var document in missing)
            {
                if (document.TaxBase == 0M && document.VatBase == 0M)
                {
                    anulledNRA.Add(document);
                }
                else
                {
                    PrintObjectData(worksheet, document, row, 1);

                    if (Azhur_Services.Azhur_Data.Any(a => a.Id == document.Id && a.VatBase == document.VatBase))
                    {
                        Color_Services.HighlightNRAObject(worksheet, row, 1, Color.Yellow);
                    }
                    row++;
                }
            }
        }
        
        public static void WrongVAT(ExcelWorksheet worksheet)
        {
            //Вземаме само единични документи от НАП, които не са сторнирани
            var matchingDocuments = NRA_Data
                .GroupJoin(
                    Azhur_Services.Azhur_Data,
                    n => new { n.Id, n.DocumentNum },
                    a => new { a.Id, a.DocumentNum },
                    (n, a) => new
                    {
                        NRA_document = n,
                        Azhur_documents = a.ToList()
                    })
                    .Where(group => group.Azhur_documents.Count == 1)
                    .GroupBy(group => group.NRA_document.DocumentNum)
                    .Where(group => group.Count() == 1)
                    .SelectMany(group => group)
                    .ToList();

            int row = 2;
            worksheet.Cells[1, 20].Value = "Разлика в ДДС";

            foreach (var document in matchingDocuments)
            {
                if (Math.Abs(document.NRA_document.VatBase - document.Azhur_documents.First().VatBase) > 0.5M &&
                        document.Azhur_documents.First().VatBase != 0M)
                {
                    PrintObjectData(worksheet, document.NRA_document, row, 1);
                    Azhur_Services.PrintObjectData(worksheet, document.Azhur_documents.First(), row, 10);
                    worksheet.Cells[row, 20].Value = document.NRA_document.VatBase - document.Azhur_documents.First().VatBase;
                    row++;
                }
            }

            //Вземане на документи, които фигурират 2 пъти в НАП
            var doubledDocuments = Azhur_Services.Azhur_Data
                .GroupJoin(
                    NRA_Data,
                    a => new { a.Id, a.DocumentNum },
                    n => new { n.Id, n.DocumentNum },
                    (a, n) => new
                    {
                        Azhur_document = a,
                        NRA_documents = n.ToList()
                    })
                    .Where(group => group.NRA_documents.Count > 1)
                    .GroupBy(group => group.Azhur_document.DocumentNum)
                    .Where(group => group.Count() == 1)
                    .SelectMany(group => group)
                    .ToList();

            foreach (var document in doubledDocuments)
            {
                var sumOfDoubledNRA = document.NRA_documents.Sum(n => n.VatBase);
                if (Math.Abs(document.Azhur_document.VatBase - sumOfDoubledNRA) > 0.5M)
                {
                    Azhur_Services.PrintObjectData(worksheet, document.Azhur_document, row, 10);
                    for (int nraRow = row; nraRow < document.NRA_documents.Count + row; nraRow++)
                    {
                        PrintObjectData(worksheet, document.NRA_documents[nraRow - row], nraRow, 1);
                    }
                    row += document.NRA_documents.Count;
                }
            }
        }
        
        public static void CancelledDocuments(ExcelWorksheet worksheet)
        {
            var cancelledDocuments = NRA_Data
                .GroupJoin(
                    Azhur_Services.Azhur_Data,
                    n => new { n.Id, n.DocumentNum },
                    a => new { a.Id, a.DocumentNum },
                    (n, a) => new
                    {
                        NRA_document = n,
                        Azhur_documents = a.ToList()
                    })
                    .Where(group => group.Azhur_documents.Count > 1)
                    .GroupBy(group => group.NRA_document.DocumentNum)
                    .Where(group => group.Count() == 1)
                    .SelectMany(group => group)
                    .ToList();

            int row = 2;
            worksheet.Cells[1, 20].Value = "Разлика в ДДС";

            foreach (var document in cancelledDocuments)
            {
                var vatSumOfCancelled = document.Azhur_documents.Sum(a => a.VatBase);
                var vatDifference = document.NRA_document.VatBase - vatSumOfCancelled;

                if (Math.Abs(vatDifference) > 0.5M)
                {
                    Color_Services.HighlightCell(worksheet, row, 20, Color.Yellow);
                }
                worksheet.Cells[row, 20].Value = vatDifference;

                PrintObjectData(worksheet, document.NRA_document, row, 1);
                for (int azhurRow = row; azhurRow < document.Azhur_documents.Count + row; azhurRow++)
                {
                    Azhur_Services.PrintObjectData(worksheet, document.Azhur_documents[azhurRow - row], azhurRow, 10);
                }
                row += document.Azhur_documents.Count;
            }
        }

        public static void AnulledDocuments(ExcelWorksheet worksheet)
        {
            for (int row = 2; row < anulledNRA.Count + 2; row++)
            {
                PrintObjectData(worksheet, anulledNRA[row - 2], row, 1);
            }
        }

        public static void PrintObjectData(ExcelWorksheet worksheet, NRA document, int row, int col)
        {
            worksheet.Cells[row, col++].Value = document.Id;
            worksheet.Cells[row, col++].Value = document.Name;
            worksheet.Cells[row, col++].Value = document.Period;
            worksheet.Cells[row, col++].Value = document.DocumentNum;
            worksheet.Cells[row, col++].Value = document.DocumentType;
            worksheet.Cells[row, col++].Value = document.DocumentDate;
            worksheet.Cells[row, col++].Value = document.MerchType;
            worksheet.Cells[row, col++].Value = document.TaxBase;
            worksheet.Cells[row, col++].Value = document.VatBase;
        }

        public static void PrintHeader(ExcelWorksheet worksheet, int col)
        {
            worksheet.Cells[1, col++].Value = "Идент. № на доставчика ВИН";
            worksheet.Cells[1, col++].Value = "Наименование на доставчика";
            worksheet.Cells[1, col++].Value = "Период";
            worksheet.Cells[1, col++].Value = "№ на документ";
            worksheet.Cells[1, col++].Value = "Тип на документ";
            worksheet.Cells[1, col++].Value = "Дата на издаване";
            worksheet.Cells[1, col++].Value = "Предмет на доставка";
            worksheet.Cells[1, col++].Value = "0210: Сума на ДО";
            worksheet.Cells[1, col++].Value = "0220: Начислен ДДС";
        }
    }
}
