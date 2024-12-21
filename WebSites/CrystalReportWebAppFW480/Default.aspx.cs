using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Web;
using System;
using System.Web.UI;

namespace CrystalReportWebAppFW480
{
    public partial class _Default : Page
    {
        protected void Page_Init(object sender, EventArgs e)
        {
            ReportDocument rpt = new ReportDocument();
            //CrystalReportSource src = new CrystalReportSource();
            var reportFile = Server.MapPath(@"~\Reports\CommRateSchedule.rpt");
            rpt.Load(reportFile);
            CrystalReportViewer1.ReportSource = rpt;
        }

        protected void Page_Load(object sender, EventArgs e)
        {

        }
    }
}