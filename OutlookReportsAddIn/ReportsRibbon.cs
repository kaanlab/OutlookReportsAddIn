using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace OutlookReportsAddIn
{
    public partial class ReportsRibbon
    {
        private void ReportsRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void ReportOnOneDay_Click(object sender, RibbonControlEventArgs e)
        {
            ReportsForm form = new ReportsForm();
            form.Show();
        }
    }
}
