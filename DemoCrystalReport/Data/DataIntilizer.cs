using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DemoCrystalReport.Data
{
    class DataIntilizer
    {
        private SAPbobsCOM.Company oCompany;

        public DataIntilizer (SAPbobsCOM.Company company)
        {
            oCompany = company;
        }

        public void Intilizer()
        {
            UDT udt = new UDT(oCompany);
            udt.Create();

            UDF udf = new UDF(oCompany);
            udf.Create();
        }
    }
}
