using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DemoUDO.Data
{
    class DataIntilizer
    {
        public void Initialize(SAPbobsCOM.Company company)
        {
            UDT udt = new UDT(company);
            udt.Create();

            UDF udf = new UDF(company);
            udf.Create();

            UDO udo = new UDO(company);
            udo.Create();
        }
    }
}
