using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DemoCrystalReport.Forms
{
    public partial class CrystalView : Form
    {
        public CrystalView()
        {
            InitializeComponent();
        }

        private void CrystalView_Load(object sender, EventArgs e)
        {
            Reports.EmployeeReport report = new Reports.EmployeeReport();

            EmployeeService service = new EmployeeService(Data.Connection.DICompany);
            DataSet ds = service.GetEmployee();

            report.SetDataSource(ds);

            crystalReportViewer1.ReportSource = report;
            crystalReportViewer1.Refresh();
        }
    }
}
