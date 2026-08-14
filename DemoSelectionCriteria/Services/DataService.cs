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

        public DataSet GetDetail(
        string fromDate,
        string toDate,
        string account,
        string cardCode,
        string cardType = "") // "C" = Customer, "S" = Vendor
        {
            string accountEsc = account.Replace("'", "''");
            string cardCodeEsc = cardCode.Replace("'", "''");
            string cardTypeEsc = cardType.Replace("'", "''");
            // Dùng riêng cho subquery Opening Balance (không có bảng con trung gian)
            var sb = new System.Text.StringBuilder();

            sb.AppendLine("SELECT");
            sb.AppendLine("    T1.Account,");
            sb.AppendLine("    T2.AcctName,");
            sb.AppendLine("    T1.ShortName AS CardCode,");
            sb.AppendLine("    T3.CardName,");
            sb.AppendLine("    T0.RefDate,");
            sb.AppendLine("    T0.TaxDate,");
            sb.AppendLine("    T0.TransId,");
            sb.AppendLine("    T0.Ref1,");
            sb.AppendLine("    T0.Memo,");
            sb.AppendLine("    T1b.Debit AS Credit,");
            sb.AppendLine("    T1b.Credit AS Debit,");
            sb.AppendLine("    T1b.Account AS ContraAccount,");
            sb.AppendLine("    CAST(NULL AS NVARCHAR(100)) AS TransferTerm,"); 
            sb.AppendLine("    ISNULL(OB.OpenDebit, 0) AS OpenDebit,");
            sb.AppendLine("    ISNULL(OB.OpenCredit, 0) AS OpenCredit");
            sb.AppendLine("FROM OJDT T0");
            sb.AppendLine("INNER JOIN JDT1 T1");
            sb.AppendLine("    ON T0.TransId = T1.TransId");
            sb.AppendLine("   AND T1.Line_ID = 0");
            sb.AppendLine("INNER JOIN JDT1 T1b");
            sb.AppendLine("    ON T1b.TransId = T1.TransId");
            sb.AppendLine("   AND T1b.Line_ID <> 0");
            sb.AppendLine("   AND T1b.Account <> T1.Account");
            sb.AppendLine("   AND (T1b.Debit <> 0 OR T1b.Credit <> 0)");
            sb.AppendLine("INNER JOIN OCRD T3");
            sb.AppendLine("    ON T1.ShortName = T3.CardCode");
            sb.AppendLine("INNER JOIN OACT T2");
            sb.AppendLine("    ON T1.Account = T2.AcctCode");
            sb.AppendLine("LEFT JOIN");
            sb.AppendLine("(");
            sb.AppendLine("    SELECT");
            sb.AppendLine("        T1.Account,");
            sb.AppendLine("        T1.ShortName AS CardCode,");
            sb.AppendLine("        ISNULL(SUM(T1.Debit), 0) AS OpenDebit,");
            sb.AppendLine("        ISNULL(SUM(T1.Credit), 0) AS OpenCredit");
            sb.AppendLine("    FROM OJDT T0");
            sb.AppendLine("    INNER JOIN JDT1 T1");
            sb.AppendLine("        ON T0.TransId = T1.TransId");
            sb.AppendLine("    INNER JOIN OCRD T3");
            sb.AppendLine("        ON T1.ShortName = T3.CardCode");
            sb.AppendLine($"    WHERE T0.RefDate < '{fromDate}'");
            if (!string.IsNullOrWhiteSpace(cardTypeEsc))
                sb.AppendLine($"      AND T3.CardType = '{cardTypeEsc}'");
            sb.AppendLine("    GROUP BY");
            sb.AppendLine("        T1.Account,");
            sb.AppendLine("        T1.ShortName");
            sb.AppendLine(") OB");
            sb.AppendLine("    ON T1.Account = OB.Account");
            sb.AppendLine("   AND T1.ShortName = OB.CardCode");
            sb.AppendLine($"WHERE T0.RefDate >= '{fromDate}'");
            sb.AppendLine($"  AND T0.RefDate < '{toDate}'");
            if (!string.IsNullOrWhiteSpace(accountEsc))
                sb.AppendLine($"  AND T1.Account = '{accountEsc}'");
            if (!string.IsNullOrWhiteSpace(cardCodeEsc))
                sb.AppendLine($"  AND T1.ShortName = '{cardCodeEsc}'");
            if (!string.IsNullOrWhiteSpace(cardTypeEsc))
                sb.AppendLine($"  AND T3.CardType = '{cardTypeEsc}'");
            sb.AppendLine("ORDER BY");
            sb.AppendLine("    T1.Account,");
            sb.AppendLine("    T1.ShortName,");
            sb.AppendLine("    T0.RefDate,");
            sb.AppendLine("    T0.TransId,");
            sb.AppendLine("    T1b.Line_ID");

            string sql = sb.ToString();
            return ExecuteQuery(sql);
        }

        private DataSet ExecuteQuery(string sql)
        {
            Recordset rs = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
            rs.DoQuery(sql);

            DataTable dt = new DataTable("LedgerDetail");
            dt.Columns.Add("Account", typeof(string)); // Ledger Account
            dt.Columns.Add("AcctName", typeof(string)); // Ledger Name
            dt.Columns.Add("CardCode", typeof(string)); // Customer/Vendor Code
            dt.Columns.Add("CardName", typeof(string)); // Customer/Vendor Name
            dt.Columns.Add("RefDate", typeof(DateTime)); // Ngày ghi sổ
            dt.Columns.Add("TransId", typeof(string)); // Số chứng từ (dùng khi Ref1 trống)
            dt.Columns.Add("Ref1", typeof(string)); // Số hiệu chứng từ (ưu tiên)
            dt.Columns.Add("TaxDate", typeof(DateTime)); // Ngày chứng từ
            dt.Columns.Add("Memo", typeof(string)); // Diễn giải
            dt.Columns.Add("ContraAccount", typeof(string)); // Tài khoản đối ứng
            dt.Columns.Add("TransferTerm", typeof(string)); // Thoi han chuyen khoan
            dt.Columns.Add("Debit", typeof(decimal)); // Số phát sinh nợ
            dt.Columns.Add("Credit", typeof(decimal)); // Số phát sinh có
            dt.Columns.Add("OpenDebit", typeof(decimal)); // Số dư nợ đầu kỳ
            dt.Columns.Add("OpenCredit", typeof(decimal)); // Số dư có đầu kỳ

            if (!rs.EoF)
            {
                rs.MoveFirst();
                while (!rs.EoF)
                {
                    DataRow row = dt.NewRow();
                    row["Account"] = rs.Fields.Item("Account").Value ?? "";
                    row["AcctName"] = rs.Fields.Item("AcctName").Value ?? "";
                    row["CardCode"] = rs.Fields.Item("CardCode").Value ?? "";
                    row["CardName"] = rs.Fields.Item("CardName").Value ?? "";
                    row["RefDate"] = Convert.ToDateTime(rs.Fields.Item("RefDate").Value);
                    row["TransId"] = rs.Fields.Item("TransId").Value ?? "";
                    row["Ref1"] = rs.Fields.Item("Ref1").Value ?? "";
                    row["TaxDate"] = Convert.ToDateTime(rs.Fields.Item("TaxDate").Value);
                    row["Memo"] = rs.Fields.Item("Memo").Value ?? "";
                    row["ContraAccount"] = rs.Fields.Item("ContraAccount").Value ?? "";
                    row["TransferTerm"] = rs.Fields.Item("TransferTerm").Value ?? "0";
                    row["Debit"] = Convert.ToDecimal(rs.Fields.Item("Debit").Value);
                    row["Credit"] = Convert.ToDecimal(rs.Fields.Item("Credit").Value);
                    row["OpenDebit"] = Convert.ToDecimal(rs.Fields.Item("OpenDebit").Value);
                    row["OpenCredit"] = Convert.ToDecimal(rs.Fields.Item("OpenCredit").Value);
                    dt.Rows.Add(row);
                    rs.MoveNext();
                }
            }

            DataSet ds = new DataSet("PaymentData");
            ds.Tables.Add(dt);

            return ds;
        }

        // Get Data
        public DataSet GetTable(string fromDate, string toDate, string account, string cardCode, string currency, string cardType = "")
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
                                INNER JOIN OCRD T3
                                    ON T1.ShortName = T3.CardCode
                                WHERE 
                                    T0.RefDate >= '{fromDate}'
                                    AND T0.RefDate <= '{toDate}'";

            if (!string.IsNullOrEmpty(account))
                sql += $" AND T1.Account = '{account}'";

            if (!string.IsNullOrEmpty(cardCode))
                sql += $" AND T1.ShortName = '{cardCode}'";

            if (!string.IsNullOrEmpty(cardType))
                sql += $" AND T3.CardType = '{cardType}'";

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
                decimal rate = currency.Trim() == "VND" ? 22000m : 1m;
                decimal debit = Convert.ToDecimal(rs.Fields.Item("Debit").Value) * rate;
                decimal credit = Convert.ToDecimal(rs.Fields.Item("Credit").Value) * rate;

                dt.Rows.Add(
                    rs.Fields.Item("PostingDate").Value,
                    rs.Fields.Item("DocNum").Value,
                    rs.Fields.Item("DocDate").Value,
                    rs.Fields.Item("Description").Value,
                    rs.Fields.Item("ContraAccount").Value,
                    rs.Fields.Item("DiscountTerm").Value == DBNull.Value ? "" : rs.Fields.Item("DiscountTerm").Value.ToString(),
                    debit,
                    credit);

                rs.MoveNext();
            }
            DataSet ds = new DataSet("PaymentData");
            ds.Tables.Add(dt);

            return ds;
        }

        // Get Opening Balance
        public DataTable GetOpeningBalance(string fromDate, string account, string cardCode, string currency, string cardType = "")
        {
            string sql = $@"SELECT
                                ISNULL(SUM(T1.Debit),0) AS OpeningDebit,
                                ISNULL(SUM(T1.Credit),0) AS OpeningCredit
                            FROM OJDT T0
                            INNER JOIN JDT1 T1
                                ON T0.TransId = T1.TransId
                            INNER JOIN OCRD T3
                                ON T1.ShortName = T3.CardCode
                            WHERE
                                T0.RefDate < '{fromDate}'";

            if (!string.IsNullOrEmpty(account))
                sql += $" AND T1.Account='{account}'";

            if (!string.IsNullOrEmpty(cardCode))
                sql += $" AND T1.ShortName='{cardCode}'";

            if (!string.IsNullOrEmpty(cardType))
                sql += $" AND T3.CardType='{cardType}'";

            Recordset rs = (SAPbobsCOM.Recordset)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rs.DoQuery(sql);

            DataTable dt = new DataTable("OpeningBalance");

            dt.Columns.Add("OpeningDebit");
            dt.Columns.Add("OpeningCredit");

            decimal rate = currency.Trim() == "VND" ? 22000m : 1m;
            decimal debit = Convert.ToDecimal(rs.Fields.Item("OpeningDebit").Value) * rate;
            decimal credit = Convert.ToDecimal(rs.Fields.Item("OpeningCredit").Value) * rate;

            dt.Rows.Add(
                debit,
                credit
            );
            return dt;
        }
    }
}
