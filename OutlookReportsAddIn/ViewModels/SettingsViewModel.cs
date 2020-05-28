using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using OutlookReportsAddIn.Helpers;

namespace OutlookReportsAddIn.ViewModels
{
    public class SettingsViewModel : BaseViewModel
    {
        public static string AssemblyVersion { get { return Assembly.GetExecutingAssembly().GetName().Version.ToString(); } }
        public static string AssemblyCopyright { get { return AssemblyAttributeHelper.GetExecutingAssemblyAttribute<AssemblyCopyrightAttribute>(a => a.Copyright); } }
        public static string AssemblyCompany { get { return AssemblyAttributeHelper.GetExecutingAssemblyAttribute<AssemblyCompanyAttribute>(a => a.Company); } }
        public static string WindowTitle { get => "Настройки"; }

        private string _mailAddress = Properties.Settings.Default.MailAddress;
        public string MailAddress
        {
            get => _mailAddress;
            set
            {
                _mailAddress = value;
                OnPropertyChanged("MailAddress");
            }
        }

        private string _templatePath = Properties.Settings.Default.TemplatePath;
        public string TemplatePath
        {
            get => _templatePath;
            set
            {
                _templatePath = value;
                OnPropertyChanged("TempalatePath");
            }
        }

        private bool _isTemplatePathExsist = File.Exists(Properties.Settings.Default.TemplatePath);
        public bool IsTemplatePathExsist
        {
            get => _isTemplatePathExsist;
            set
            {
                _isTemplatePathExsist = value;
                OnPropertyChanged("IsTemplatePathExsist");
            }
        }

        public ICommand SetTemplatePathCommand { get; }
        public ICommand SaveCommand { get; }

        public SettingsViewModel()
        {
            SetTemplatePathCommand = new RelayCommand(SetTemplate);
            SaveCommand = new RelayCommand(Save);
        }

        private void Save()
        {
            Properties.Settings.Default.MailAddress = MailAddress;
            Properties.Settings.Default.Save();
        }

        private void SetTemplate()
        {
            // Configure save file dialog box
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.FileName = "Template"; // Default file name
            dlg.DefaultExt = ".dotx"; // Default file extension
            dlg.Filter = "Шаблон Word (.dotx)|*.dotx"; // Filter files by extension

            // Show save file dialog box
            Nullable<bool> result = dlg.ShowDialog();

            // Process save file dialog box results
            if (result == true)
            {
                Properties.Settings.Default.TemplatePath = dlg.FileName;
                TemplatePath = Properties.Settings.Default.TemplatePath;
                IsTemplatePathExsist = true; // update image     
            }
        }
    }
}
