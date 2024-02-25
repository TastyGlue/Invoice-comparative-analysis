using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Invoice_Demo_Ver_1._0.Models
{
    public class NRA
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Period { get; set; }
        public int DocumentType { get; set; }
        public long DocumentNum { get; set; }
        public DateOnly DocumentDate { get; set; }
        public string MerchType { get; set; }
        public Decimal TaxBase { get; set; }
        public Decimal VAT_Base { get; set; }
        public Decimal TaxBase20 { get; set;}
        public Decimal VAT_Base20 { get; set; }
        public Decimal TaxBase9 { get; set; }
        public Decimal VAT_Base9 { get; set; }
        public Decimal TaxBase0 { get; set; }
        public int VAT_Base0 { get; set; } = 0;

        

        public NRA(string id, string name, string period, int documentType, long documentNum, DateOnly documentDate, string merchType, decimal taxBase, decimal vAT_Base, decimal taxBase20, decimal vAT_Base20, decimal taxBase9, decimal vAT_Base9, decimal taxBase0)
        {
            Id = id;
            Name = name;
            Period = period;
            DocumentType = documentType;
            DocumentNum = documentNum;
            DocumentDate = documentDate;
            MerchType = merchType;
            TaxBase = taxBase;
            VAT_Base = vAT_Base;
            TaxBase20 = taxBase20;
            VAT_Base20 = vAT_Base20;
            TaxBase9 = taxBase9;
            VAT_Base9 = vAT_Base9;
            TaxBase0 = taxBase0;
        }
    }
}
