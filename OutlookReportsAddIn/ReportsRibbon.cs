using Microsoft.Office.Tools.Ribbon;
using OutlookReportsAddIn.Views;

namespace OutlookReportsAddIn
{
    public partial class ReportsRibbon
    {
        private void ReportsRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void OneDayReport_Click(object sender, RibbonControlEventArgs e)
        {
            var view = new OneDayReportView();
            view.Show();
        }

        private void ReportForPeriod_Click(object sender, RibbonControlEventArgs e)
        {
            var view = new ReportForPeriodView();
            view.Show();
        }

        private void AboutButton_Click(object sender, RibbonControlEventArgs e)
        {
            var view = new AboutView();
            view.Show();
        }
    }
}
