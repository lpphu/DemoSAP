using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DemoSelectionCriteria.Services
{
    class DataService
    {
        private Company company;
        public DataService()
        {
            Data.Connection.Initialize();
            company = Data.Connection.DICompany;

            if (company == null)
            {
                throw new Exception("DICompany is null");
            }

            if (!company.Connected)
            {
                throw new Exception("DICompany is not connected");
            }
        }

        // Get Data
        public DataSet GetTable(string fromDate, string toDate, string account, string cardCode)
        {
            string sql = $@"SELECT
                                    T0.RefDate        AS PostingDate,
                                    T0.TransId        AS DocNum,
                                    T0.TaxDate        AS DocDate,
                                    T1.LineMemo       AS Description,

                                    -- TK đối ứng
                                    STUFF(
                                        (
                                            SELECT ',' + T2.Account
                                            FROM JDT1 T2
                                            WHERE T2.TransId = T1.TransId
                                            AND T2.Account <> T1.Account
                                            FOR XML PATH('')), 1, 1, '')    AS ContraAccount,

                                    NULL AS DiscountTerm,
                                    T1.Debit          AS Debit,
                                    T1.Credit         AS Credit
                                FROM OJDT T0
                                INNER JOIN JDT1 T1 
                                    ON T0.TransId = T1.TransId
                                WHERE 
                                    T0.RefDate >= '{fromDate}'
                                    AND T0.RefDate <= '{toDate}'";

            if (!string.IsNullOrEmpty(account))
                sql += $" AND T1.Account='{account}'";

            if (!string.IsNullOrEmpty(cardCode))
                sql += $" AND T1.ShortName='{cardCode}'";

            Recordset rs = (SAPbobsCOM.Recordset)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rs.DoQuery(sql);

            DataTable dt = new DataTable("PaymentReport");

            dt.Columns.Add("PostingDate", typeof(DateTime));
            dt.Columns.Add("DocNum");
            dt.Columns.Add("DocDate", typeof(DateTime));
            dt.Columns.Add("Description");
            dt.Columns.Add("ContraAccount");
            dt.Columns.Add("DiscountTerm");
            dt.Columns.Add("Debit", typeof(decimal));
            dt.Columns.Add("Credit", typeof(decimal));

            while (!rs.EoF)
            {
                dt.Rows.Add(

                rs.Fields.Item("PostingDate").Value,
                rs.Fields.Item("DocNum").Value,
                rs.Fields.Item("DocDate").Value,
                rs.Fields.Item("Description").Value,
                rs.Fields.Item("ContraAccount").Value,
                rs.Fields.Item("DiscountTerm").Value == DBNull.Value ? "" : rs.Fields.Item("DiscountTerm").Value.ToString(),
                rs.Fields.Item("Debit").Value,
                rs.Fields.Item("Credit").Value

                );

                rs.MoveNext();
            }
            DataSet ds = new DataSet("PaymentData");
            ds.Tables.Add(dt);

            return ds;
        }

        // Get Opening Balance
        public DataTable GetOpeningBalance(string fromDate, string account, string cardCode)
        {
            string sql = $@"SELECT
                                ISNULL(SUM(T1.Debit),0) AS OpeningDebit,
                                ISNULL(SUM(T1.Credit),0) AS OpeningCredit
                            FROM OJDT T0
                            INNER JOIN JDT1 T1
                                ON T0.TransId = T1.TransId
                            WHERE
                                T0.RefDate < '{fromDate}'";

            if (!string.IsNullOrEmpty(account))
                sql += $" AND T1.Account='{account}'";

            if (!string.IsNullOrEmpty(cardCode))
                sql += $" AND T1.ShortName='{cardCode}'";

            Recordset rs = (SAPbobsCOM.Recordset)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rs.DoQuery(sql);

            DataTable dt = new DataTable("OpeningBalance");

            dt.Columns.Add("OpeningDebit");
            dt.Columns.Add("OpeningCredit");

            dt.Rows.Add(
                rs.Fields.Item("OpeningDebit").Value,
                rs.Fields.Item("OpeningCredit").Value
            );

            return dt;
        }
    }
}
