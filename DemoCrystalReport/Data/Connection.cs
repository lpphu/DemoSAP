using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DemoCrystalReport.Data
{
    class Connection
    {
        /// <summary>
        /// UI API Application
        /// </summary>
        public static SAPbouiCOM.Application UIApplication { get; private set; }

        /// <summary>
        /// DI API Company
        /// </summary>
        public static Company DICompany { get; private set; }

        public static void Initialize()
        {
            // UI API
            UIApplication = SAPbouiCOM.Framework.Application.SBO_Application;

            // DI API
            DICompany = (Company)UIApplication.Company.GetDICompany();
        }
    }
}
