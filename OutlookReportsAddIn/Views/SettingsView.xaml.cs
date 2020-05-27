using OutlookReportsAddIn.ViewModels;
using System.Windows;


namespace OutlookReportsAddIn.Views
{
    /// <summary>
    /// Interaction logic for AboutView.xaml
    /// </summary>
    public partial class SettingsView : Window
    {
        public SettingsView()
        {
            this.DataContext = new SettingsViewModel();
            InitializeComponent();
        }
    }
}
