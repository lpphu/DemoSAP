using CrystalDecisions.CrystalReports.Engine;
using DemoSelectionCriteria.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DemoSelectionCriteria.Helpers
{
    class ReportViewerHelper
    {
        public static void ShowReport(ReportDocument rpt)
        {
            Thread staThread = new Thread(() =>
            {
                try
                {
                    ReportViewerForm viewer = new ReportViewerForm(rpt);
                    viewer.ShowDialog();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            });
            staThread.SetApartmentState(ApartmentState.STA);
            staThread.Start();
            staThread.Join();
        }
    }
}
