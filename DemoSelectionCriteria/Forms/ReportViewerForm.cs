using CrystalDecisions.CrystalReports.Engine;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DemoSelectionCriteria.Forms
{
    public partial class ReportViewerForm : Form
    {
        public ReportViewerForm(ReportDocument rpt)
        {
            InitializeComponent();

            crystalReportViewer1.ReportSource = rpt;
            crystalReportViewer1.Refresh();
        }

        private void ReportViewerForm_Load(object sender, EventArgs e)
        {

        }
    }
}
