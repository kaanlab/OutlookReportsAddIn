using System.Windows;

namespace OutlookReportsAddIn
{
    /// <summary>
    /// Interaction logic for ReportsForm.xaml
    /// </summary>
    public partial class ReportsForm : Window
    {
        public ReportsForm()
        {
            InitializeComponent();
            DataContext = new ReportViewModel();
        }       
    }
}
