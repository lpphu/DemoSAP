using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DemoUDO.Data
{
    class UDO
    {
        private SAPbobsCOM.Company oCompany;
        public UDO(SAPbobsCOM.Company company)
        {
            oCompany = company;
        }
        public void Create()
        {
            CreateObject("UDO_MDATA", "Master Data", "MDATA",
                BoUDOObjType.boud_MasterData,
                new string[]
                {
                    "Code",
                    "Name",
                    "U_Local",
                    "U_Status"
                },
                new string[]
                {
                    "Code",
                    "Name",
                });

            CreateObject("UDO_DOC", "Document", "DOCUMENT",
                BoUDOObjType.boud_Document,
                new string[]
                {
                    "DocEntry",
                    "U_Text",
                    "U_Status"
                },
                new string[]
                {
                    "U_Text",
                    "U_Status"
                });
        }

        public void CreateObject(
            string code,
            string name,
            string tableName,
            BoUDOObjType objectType,
            string[] formColumns,
            string[] findColumns,
            List<string> childTables = null)
        {
            UserObjectsMD oUDO =
                (UserObjectsMD)oCompany.GetBusinessObject(BoObjectTypes.oUserObjectsMD);

            try
            {
                if (oUDO.GetByKey(code))
                    return;

                oUDO.Code = code;
                oUDO.Name = name;
                oUDO.TableName = tableName;
                oUDO.ObjectType = objectType;

                if (childTables != null)
                {
                    foreach (string child in childTables)
                    {
                        oUDO.ChildTables.Add();
                        oUDO.ChildTables.TableName = child;
                    }
                }


                // Default Form
                oUDO.CanCreateDefaultForm = BoYesNoEnum.tYES;
                oUDO.EnableEnhancedForm = BoYesNoEnum.tYES;

                // Menu
                oUDO.MenuItem = BoYesNoEnum.tYES;
                oUDO.MenuCaption = name;
                oUDO.FatherMenuID = 43520;
                oUDO.Position = 1;

                // Permission
                oUDO.CanFind = BoYesNoEnum.tYES;
                oUDO.CanDelete = BoYesNoEnum.tYES;
                oUDO.CanCancel = BoYesNoEnum.tYES;
                oUDO.CanClose = BoYesNoEnum.tYES;
                oUDO.CanYearTransfer = BoYesNoEnum.tNO;
                oUDO.ManageSeries = BoYesNoEnum.tNO;
                oUDO.CanLog = BoYesNoEnum.tNO;

                // Form Columns
                foreach (string col in formColumns)
                {
                    oUDO.FormColumns.FormColumnAlias = col;
                    oUDO.FormColumns.FormColumnDescription = col;
                    if (col == "DocEntry" || col == "Code")
                        oUDO.FormColumns.Editable = BoYesNoEnum.tNO;
                    else
                        oUDO.FormColumns.Editable = BoYesNoEnum.tYES;
                    oUDO.FormColumns.Add();
                }

                // Find Columns
                foreach (string col in findColumns)
                {
                    oUDO.FindColumns.ColumnAlias = col;
                    oUDO.FindColumns.ColumnDescription = col;
                    oUDO.FindColumns.Add();
                }

                int ret = oUDO.Add();

                if (ret != 0)
                {
                    oCompany.GetLastError(out int errCode, out string errMsg);
                    throw new Exception($"{errCode}: {errMsg}");
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUDO);
            }
        }
    }
}
