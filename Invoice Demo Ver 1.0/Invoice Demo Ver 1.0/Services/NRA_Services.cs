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
        public static void GetTableData(ExcelWorksheet worksheet)
        {
            for (int row = 3; row <= worksheet.Dimension.End.Row; row++)
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
            string Id = worksheet.Cells[row, 1].Value.ToString();
            string Name = worksheet.Cells[row, 2].Value.ToString();
            string Period = worksheet.Cells[row, 3].Value.ToString();

            int DocumentType = Convert.ToInt32(worksheet.Cells[row, 4].Value);
            if (DocumentType == 9)
                return;
            long DocumentNum = Convert.ToInt64(worksheet.Cells[row, 5].Value);

            DateTime date = DateTime.Parse(worksheet.Cells[row, 6].Value.ToString());
            DateOnly DocumentDate = DateOnly.FromDateTime(date);

            string MerchType = worksheet.Cells[row, 7].Value.ToString();

            Decimal TaxBase = Decimal.Parse(worksheet.Cells[row, 8].Value.ToString());
            Decimal VAT_Base = Decimal.Parse(worksheet.Cells[row, 9].Value.ToString());
            Decimal TaxBase20 = Decimal.Parse(worksheet.Cells[row, 10].Value.ToString());
            Decimal VAT_Base20 = Decimal.Parse(worksheet.Cells[row, 11].Value.ToString());
            Decimal TaxBase9 = Decimal.Parse(worksheet.Cells[row, 12].Value.ToString());
            Decimal VAT_Base9 = Decimal.Parse(worksheet.Cells[row, 13].Value.ToString());
            Decimal TaxBase0 = Decimal.Parse(worksheet.Cells[row, 14].Value.ToString());


            NRA_Data.Add(new NRA(Id, Name, Period, DocumentType, DocumentNum, DocumentDate, MerchType, TaxBase, VAT_Base, TaxBase20, VAT_Base20, TaxBase9, VAT_Base9, TaxBase0));
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
            foreach (var document in NRA_Data.Where(x => x.TaxBase < 0))
            {
                if (document.VAT_Base > 0)
                {
                    document.VAT_Base = document.VAT_Base * -1;
                }

                if (document.VAT_Base20 > 0)
                {
                    document.VAT_Base20 = document.VAT_Base20 * -1;
                }

                if (document.VAT_Base9 > 0)
                {
                    document.VAT_Base9 = document.VAT_Base9 * -1;
                }
            }
        }

        public static void FixInvoiceSigns()
        {
            foreach (var document in NRA_Services.NRA_Data.Where(x => x.DocumentType == 1))
            {
                if (document.TaxBase < 0)
                {
                    document.TaxBase = document.TaxBase * -1;
                }

                if (document.VAT_Base < 0)
                {
                    document.VAT_Base = document.VAT_Base * -1;
                }

                if (document.TaxBase20 < 0)
                {
                    document.TaxBase20 = document.TaxBase20 * -1;
                }

                if (document.VAT_Base20 < 0)
                {
                    document.VAT_Base20 = document.VAT_Base20 * -1;
                }

                if (document.TaxBase9 < 0)
                {
                    document.TaxBase9 = document.TaxBase9 * -1;
                }

                if (document.VAT_Base9 < 0)
                {
                    document.VAT_Base9 = document.VAT_Base9 * -1;
                }

                if (document.TaxBase0 < 0)
                {
                    document.TaxBase0 = document.TaxBase0 * -1;
                }
            }
        }

        public static void Missing(ExcelWorksheet worksheet)
        {
            var missing = NRA_Data.Where(n => !Azhur_Services.Azhur_Data.Any(a => a.DocumentNum == n.DocumentNum)).ToList();

            int row = 2;
            foreach (var document in missing)
            {
                PrintObjectData(worksheet, document, row, 1);

                if(Azhur_Services.Azhur_Data.Any(a => a.Id == document.Id && a.VAT_Base == document.VAT_Base))
                {
                    Color_Services.HighlightNRAObject(worksheet, row, 1, Color.Yellow);
                }
                row++;
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
            worksheet.Cells[1, 27].Value = "Разлика в ДДС";

            foreach (var document in matchingDocuments)
            {
                if (Math.Abs(document.NRA_document.VAT_Base - document.Azhur_documents.First().VAT_Base) > 0.5M &&
                        document.Azhur_documents.First().VAT_Base != 0M)
                {
                    PrintObjectData(worksheet, document.NRA_document, row, 1);
                    Azhur_Services.PrintObjectData(worksheet, document.Azhur_documents.First(), row, 17);
                    worksheet.Cells[row, 27].Value = document.NRA_document.VAT_Base - document.Azhur_documents.First().VAT_Base;
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
                var sumOfDoubledNRA = document.NRA_documents.Sum(n => n.VAT_Base);
                if (Math.Abs(document.Azhur_document.VAT_Base - sumOfDoubledNRA) > 0.5M)
                {
                    Azhur_Services.PrintObjectData(worksheet, document.Azhur_document, row, 17);
                    for (int nraRow = row; nraRow < document.NRA_documents.Count() + row; nraRow++)
                    {
                        PrintObjectData(worksheet, document.NRA_documents[nraRow - row], nraRow, 1);
                    }
                    row = row + document.NRA_documents.Count();
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
            worksheet.Cells[1, 27].Value = "Разлика в ДДС";

            foreach (var document in cancelledDocuments)
            {
                var vatSumOfCancelled = document.Azhur_documents.Sum(a => a.VAT_Base);
                var vatDifference = document.NRA_document.VAT_Base - vatSumOfCancelled;

                if (Math.Abs(vatDifference) > 0.5M)
                {
                    Color_Services.HighlightCell(worksheet, row, 27, Color.Yellow);
                }
                worksheet.Cells[row, 27].Value = vatDifference;

                PrintObjectData(worksheet, document.NRA_document, row, 1);
                for (int azhurRow = row; azhurRow < document.Azhur_documents.Count() + row; azhurRow++)
                {
                    Azhur_Services.PrintObjectData(worksheet, document.Azhur_documents[azhurRow - row], azhurRow, 17);
                }
                row = row + document.Azhur_documents.Count();
            }
        }

        public static void PrintObjectData(ExcelWorksheet worksheet, NRA document, int row, int col)
        {
            worksheet.Cells[row, col++].Value = document.Id;
            worksheet.Cells[row, col++].Value = document.Name;
            worksheet.Cells[row, col++].Value = document.Period;
            worksheet.Cells[row, col++].Value = document.DocumentType;
            worksheet.Cells[row, col++].Value = document.DocumentNum;
            worksheet.Cells[row, col++].Value = document.DocumentDate;
            worksheet.Cells[row, col++].Value = document.MerchType;
            worksheet.Cells[row, col++].Value = document.TaxBase;
            worksheet.Cells[row, col++].Value = document.VAT_Base;
            worksheet.Cells[row, col++].Value = document.TaxBase20;
            worksheet.Cells[row, col++].Value = document.VAT_Base20;
            worksheet.Cells[row, col++].Value = document.TaxBase9;
            worksheet.Cells[row, col++].Value = document.VAT_Base9;
            worksheet.Cells[row, col++].Value = document.TaxBase0;
            worksheet.Cells[row, col++].Value = document.VAT_Base0;
        }

        public static void PrintHeader(ExcelWorksheet worksheet, int col)
        {
            worksheet.Cells[1, col++].Value = "Идентификационен номер на контрагента";
            worksheet.Cells[1, col++].Value = "Име на контрагента";
            worksheet.Cells[1, col++].Value = "Данъчен период";
            worksheet.Cells[1, col++].Value = "Вид на документа";
            worksheet.Cells[1, col++].Value = "Номер на документа";
            worksheet.Cells[1, col++].Value = "Дата на документа";
            worksheet.Cells[1, col++].Value = "Вид на стоката или обхват и вид на услугата";
            worksheet.Cells[1, col++].Value = "Общ размер на данъчните основи за облагане с ДДС";
            worksheet.Cells[1, col++].Value = "Всичко начислен ДДС";
            worksheet.Cells[1, col++].Value = "Данъчна основа на облагаемите доставки със ставка 20 %, вкл, доставките при условията на дистанционни продажби, с място на изпълнение на територията на страната";
            worksheet.Cells[1, col++].Value = "Начислен ДДС 20 %";
            worksheet.Cells[1, col++].Value = "ДО на облагаемите доставки съсставка 9 %";
            worksheet.Cells[1, col++].Value = "Начислен ДДС 9 %";
            worksheet.Cells[1, col++].Value = "ДО на доставките със ставка 0 % по глава трета от ЗДДС";
            worksheet.Cells[1, col++].Value = "ДО на освободени доставки и освободените ВОП";
        }
    }
}
