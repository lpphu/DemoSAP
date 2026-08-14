using CrystalDecisions.CrystalReports.Engine;
using DemoSelectionCriteria.Helpers;
using DemoSelectionCriteria.Services;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace DemoSelectionCriteria.Forms
{
    [FormAttribute("DemoSelectionCriteria.Forms.VendorForm", "Forms/VendorForm.b1f")]
    class VendorForm : UserFormBase
    {
        private Services.DataService service;
        public VendorForm()
        {
            service = new Services.DataService();
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_1").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_3").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("Item_4").Specific));
            this.EditText2.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.EditText2_ChooseFromListAfter);
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_5").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("Item_6").Specific));
            this.EditText3.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.EditText3_ChooseFromListAfter);
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_7").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_8").Specific));
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_0").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("Item_2").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_9").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_10").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private void OnCustomInitialize()
        {
            CFLFilter();
            CenterForm();
            LoadComboBox();
        }

        // LoadComboBox
        private void LoadComboBox()
        {
            ComboBox0.ValidValues.Add("USD", "USD");
            ComboBox0.ValidValues.Add("VND", "VND");

            ComboBox0.Select("USD", SAPbouiCOM.BoSearchKey.psk_ByValue);
        }

        private void CenterForm()
        {
            SAPbouiCOM.Form form = (SAPbouiCOM.Form)this.UIAPIRawForm;

            int screenWidth = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea.Width;
            int screenHeight = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea.Height;

            form.Left = (screenWidth - form.Width) / 2 - UIAPIRawForm.ClientWidth;
            form.Top = (screenHeight - form.Height) / 2 - UIAPIRawForm.ClientHeight;
        }

        // Load Report
        private void LoadReport()
        {
            string fromDate = EditText0.Value;
            string toDate = EditText1.Value;
            string account = EditText2.Value.Trim();
            string cardCode = EditText3.Value.Trim();
            string currency = ComboBox0.Value.Trim();

            if (string.IsNullOrEmpty(fromDate) || string.IsNullOrEmpty(toDate))
                return;

            DataSet ds = service.GetTable(fromDate, toDate, account, cardCode, currency, "S");
            DataTable opening = service.GetOpeningBalance(fromDate, account, cardCode, currency, "S");
            DataTable dt = ds.Tables[0];

            DataRow row = dt.NewRow();

            row["PostingDate"] = DBNull.Value;
            row["DocNum"] = "";
            row["DocDate"] = DBNull.Value;
            row["Description"] = "Số dư đầu kỳ";
            row["ContraAccount"] = "-";
            row["DiscountTerm"] = "-";

            row["Debit"] = 0m;
            row["Credit"] = 0m;

            dt.Rows.InsertAt(row, 0);

            ReportDocument rpt = new ReportDocument();

            string reportPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                            "Reports",
                                            "VendorReport.rpt");
            rpt.Load(reportPath);
            rpt.SetDataSource(ds.Tables[0]);

            DateTime fDate = DateTime.ParseExact(EditText0.Value, "yyyyMMdd", null);
            rpt.SetParameterValue("FromDate", fDate);

            DateTime tDate = DateTime.ParseExact(EditText1.Value, "yyyyMMdd", null);
            rpt.SetParameterValue("ToDate", tDate);

            rpt.SetParameterValue("LedgerAccount", account);
            rpt.SetParameterValue("Vendor", cardCode);

            rpt.SetParameterValue("Currency", currency);

            // Add 
            rpt.SetParameterValue("OpeningDebit", opening.Rows[0]["OpeningDebit"]);
            rpt.SetParameterValue("OpeningCredit", opening.Rows[0]["OpeningCredit"]);

            ReportViewerHelper.ShowReport(rpt);
        }

        private void CFLFilter()
        {
            SAPbouiCOM.ChooseFromList cfl = UIAPIRawForm.ChooseFromLists.Item("CFL_VA");

            SAPbouiCOM.Conditions conditions = new SAPbouiCOM.Conditions();
            SAPbouiCOM.Condition condition = conditions.Add();

            condition.Alias = "CardType";
            condition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            condition.CondVal = "S";

            cfl.SetConditions(conditions);
        }

        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.Button Button0;

        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            //throw new System.NotImplementedException();
            LoadReport();
        }

        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.EditText EditText1;

        private void EditText2_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            SAPbouiCOM.ISBOChooseFromListEventArg cflEvent = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;

            SAPbouiCOM.DataTable oDataTable = cflEvent.SelectedObjects;
            string code = oDataTable.GetValue("AcctCode", 0).ToString();

            this.EditText2.Value = code;
        }

        private void EditText3_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            SAPbouiCOM.ISBOChooseFromListEventArg cflEvent = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;

            SAPbouiCOM.DataTable oDataTable = cflEvent.SelectedObjects;
            string code = oDataTable.GetValue("CardCode", 0).ToString();

            this.EditText3.Value = code;
        }

        private SAPbouiCOM.StaticText StaticText4;
        private SAPbouiCOM.ComboBox ComboBox0;
    }
}
