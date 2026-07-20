using DemoCrystalReport.Data;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DemoCrystalReport
{
    class EmployeeService
    {
        private Company company;


        public EmployeeService(Company company)
        {
            this.company = company;
        }



        public EmployeeDS GetEmployee(string code = "")
        {

            EmployeeDS ds = new EmployeeDS();


            Recordset rs =
                (Recordset)company
                .GetBusinessObject(
                    BoObjectTypes.BoRecordset);



            string sql = @"SELECT
                    Code,
                    Name,
                    U_Location,
                    U_Status

                FROM 
                    [@EMPLOYEE]
                WHERE 1=1
            ";



            if (!string.IsNullOrEmpty(code))
            {
                sql +=
                $" AND Code='{code}'";
            }

            rs.DoQuery(sql);

            while (!rs.EoF)
            {
                DataRow row = ds.Employee.NewRow();

                row["Code"] = rs.Fields.Item("Code").Value.ToString();
                row["Name"] = rs.Fields.Item("Name").Value.ToString();
                row["Location"] = rs.Fields.Item("U_Location").Value.ToString();
                row["Status"] = rs.Fields.Item("U_Status").Value.ToString();

                ds.Employee.Rows.Add(row);
                rs.MoveNext();
            }
            return ds;
        }
    }
}
