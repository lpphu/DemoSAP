using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DemoUDO.Data
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
            // Mster Data
            CreateField("MDATA", "Local", "Local", BoFieldTypes.db_Alpha, 50);
            CreateField("MDATA", "Status", "Status", BoFieldTypes.db_Alpha, 1,
                new Dictionary<string, string>
                {
                    {"A", "Active"},
                    {"I", "Inactive"},
                    {"L", "Leave"}
                },
                "A");

            // Document 
            CreateField("DOCUMENT", "Text", "Text", BoFieldTypes.db_Alpha, 20);
            CreateField("DOCUMENT", "Status", "Status", BoFieldTypes.db_Alpha, 1,
                new Dictionary<string, string>
                {
                    {"A", "Active"},
                    {"I", "Inactive"},
                    {"L", "Leave"}
                },
                "A");

            //// Docuemnt Lines
            //CreateField("DOCUMENT_L", "AssetCode", "Asset", BoFieldTypes.db_Alpha,20);


            //CreateField(
            //    "UDO_ASSETASSIGNMENT_L",
            //    "ReturnDate",
            //    "Return Date",
            //    BoFieldTypes.db_Date
            //);


        }
        private void CreateField(
            string tableName,
            string fieldName,
            string description,
            BoFieldTypes type,
            int size = 0,
            Dictionary<string, string> validValues = null,
            string defaultValue = "")
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

            // ComboBox / CheckBox
            if (validValues != null)
            {
                foreach (var item in validValues)
                {
                    oUDF.ValidValues.Value = item.Key;
                    oUDF.ValidValues.Description = item.Value;
                    oUDF.ValidValues.Add();
                }

                oUDF.DefaultValue = defaultValue;
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
