using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Invoice_Demo_Ver_1._0.Models
{
    public class Azhur
    {
        public string Id { get; set; }
        public string DocumentType { get; set; }
        public string DocumentNum { get; set; }
        public DateOnly DocumentDate { get; set; }
        public string Name { get; set; }
        public Decimal NoTax { get; set; }
        public Decimal TaxBase { get; set; }
        public Decimal VatBase { get; set; }
        public string Article {get; set; }

        

        public Azhur(string id, string documentType, string documentNum, DateOnly documentDate, string name, decimal noTax, decimal taxBase, decimal vAT_Base, string article)
        {
            Id = id;
            DocumentType = documentType;
            DocumentNum = documentNum;
            DocumentDate = documentDate;
            Name = name;
            NoTax = noTax;
            TaxBase = taxBase;
            VatBase = vAT_Base;
            Article = article;
        }
    }
}
