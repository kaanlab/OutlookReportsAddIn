using OutlookReportsAddIn.ViewModels;
using System.Windows;


namespace OutlookReportsAddIn.Views
{
    /// <summary>
    /// Interaction logic for AboutView.xaml
    /// </summary>
    public partial class AboutView : Window
    {
        public AboutView()
        {
            this.DataContext = new AboutViewModel();
            InitializeComponent();
        }
    }
}
