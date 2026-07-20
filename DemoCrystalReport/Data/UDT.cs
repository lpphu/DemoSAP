using SAPbobsCOM;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DemoCrystalReport.Data
{
    class UDT
    {
        private SAPbobsCOM.Company oCompany;

        public UDT(SAPbobsCOM.Company company)
        {
            oCompany = company;
        }

        public void Create()
        {
            CreateTable("EMPLOYEE", "Employee", BoUTBTableType.bott_NoObject);
        }

        public void CreateTable(string name, string des, BoUTBTableType tableType)
        {


            UserTablesMD table = (UserTablesMD)oCompany.GetBusinessObject(BoObjectTypes.oUserTables);
            if (table.GetByKey(name))
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(table);
                return;
            }
            table.TableName = name;
            table.TableDescription = des;
            table.TableType = tableType;

            int result = table.Add();

            if (result != 0)
            {
                Application.SBO_Application.MessageBox(oCompany.GetLastErrorDescription());
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(table);
        }
    }
}
