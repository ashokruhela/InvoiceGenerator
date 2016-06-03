using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InvoiceGenerator
{
    class CompanyInfo
    {
        public readonly Company company = Company.SHOPON4U;
        public CompanyInfo(Company companyName)
        {
            company = companyName;
        }

        public string Name { get; set; }

        public string Tin { get; set; }

        public string Email { get; set; }

        public string Logo { get; set; }

        public string Website { get; set; }

        public void Initialize()
        {
            switch (company)
            {
                case Company.SHOPON4U:
                    Email = @"info@shopon4u.com";
                    Website = "www.shopon4u.com";
                    break;
                case Company.STYELOBY:
                    break;
                case Company.CRUSEBEEN:
                    break;
                default:
                    break;
            }
        }
    }
}
