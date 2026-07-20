using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DemoCrystalReport.Data
{
    class UDF
    {
        private SAPbobsCOM.Company oCompany;

        public UDF(SAPbobsCOM.Company company)
        {
            oCompany = company;
        }

        public void Create()
        {
            CreateField("EMPLOYEE", "Location", "Location", BoFieldTypes.db_Alpha, 20);
            CreateField("EMPLOYEE", "Status", "Status", BoFieldTypes.db_Alpha, 1);
        }
        private void CreateField(
            string tableName,
            string fieldName,
            string description,
            BoFieldTypes type,
            int size = 0)
        {
            UserFieldsMD oUDF = (UserFieldsMD)oCompany.GetBusinessObject(BoObjectTypes.oUserFields);

            oUDF.TableName = tableName;
            oUDF.Name = fieldName;
            oUDF.Description = description;
            oUDF.Type = type;
            if (type == BoFieldTypes.db_Alpha)
            {
                oUDF.Size = size;
            }
            else if (type == BoFieldTypes.db_Float)
            {
                oUDF.EditSize = 11;
            }

            int ret = oUDF.Add();

            if (ret != 0)
            {
                int errCode;
                string errMsg;
                oCompany.GetLastError(out errCode, out errMsg);
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUDF);
        }
    }
}
