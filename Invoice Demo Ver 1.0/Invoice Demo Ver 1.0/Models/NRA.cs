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
        public string DocumentNum { get; set; }
        public DateOnly DocumentDate { get; set; }
        public string MerchType { get; set; }
        public Decimal TaxBase { get; set; }
        public Decimal VatBase { get; set; }

        

        public NRA(string id, string name, string period, int documentType, string documentNum, DateOnly documentDate, string merchType, decimal taxBase, decimal vatBase)
        {
            Id = id;
            Name = name;
            Period = period;
            DocumentType = documentType;
            DocumentNum = documentNum;
            DocumentDate = documentDate;
            MerchType = merchType;
            TaxBase = taxBase;
            VatBase = vatBase;
        }
    }
}
