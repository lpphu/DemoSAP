using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Xml;

namespace DemoCFL
{
    [FormAttribute("DemoCFL.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {
        public static SAPbobsCOM.Company oCompanyStatic;
        public SAPbobsCOM.Recordset oRecordset;
        public SAPbobsCOM.Company oCompany;
        public SAPbouiCOM.Form oForm;
        public SAPbouiCOM.Application oApplication;

        private string _lastLoadedPO = "";
        public Form1()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_0").Specific));
            this.EditText0.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.EditText0_ChooseFromListAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.EditText EditText0;

        private void OnCustomInitialize()
        {

        }

        private void EditText0_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.ISBOChooseFromListEventArg cflEvent = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;

            SAPbouiCOM.DataTable oDataTable = cflEvent.SelectedObjects;
            string selectedDocEntry = oDataTable.GetValue("DocEntry", 0).ToString();

            this.EditText0.Value = selectedDocEntry;
        }
    }
}