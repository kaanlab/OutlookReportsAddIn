using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text;
using System.Windows.Input;
using System.Xml;
using System.Xml.Serialization;

namespace OutlookReportsAddIn
{
    public class ReportViewModel : BaseViewModel
    {
        public string Title { get; set; } = "Журнал принятой/отправленной корреспонденции";
        public string AddInVersion { get; private set; }
        public string AddInCompany { get; private set; }
        public string AddInCopyright { get; private set; }
        public ObservableCollection<Mail> ItemsCollection { get; set; }
        public DateTime SetDate { get; set; } = DateTime.Now;

        public string TemplatePath { get; set; } = Properties.Settings.Default.TemplatePath;

        public bool HasItems { get { return ItemsCollection?.Count > 0; } }

        public ICommand FetchItemsCommand { get; set; }
        public ICommand CreateReportCommand { get; set; }

        public ICommand SetTemplatePathCommand { get; set; }

        public ReportViewModel()
        {

            //var version = "";
            //version..GetVersion();
            AddInVersion = "ver.: " + AssemblyInfoHelper.Version.ToString();
            AddInCompany = AssemblyInfoHelper.Company.ToString();
            AddInCopyright = AssemblyInfoHelper.Copyright.ToString();

            FetchItemsCommand = new RelayCommand(FetchItems);
            CreateReportCommand = new RelayCommand(CreateReport);
            SetTemplatePathCommand = new RelayCommand(SetTemplate);

        }

        private void FetchItems()
        {
            var inboxMails = MailHelper.SearchInFolder(SetDate, OlDefaultFolders.olFolderInbox);
            var sentMails = MailHelper.SearchInFolder(SetDate, OlDefaultFolders.olFolderSentMail);

            var itemsList = new List<Mail>();
            itemsList.AddRange(inboxMails);
            itemsList.AddRange(sentMails);

            ItemsCollection = new ObservableCollection<Mail>(itemsList);
        }

        private void CreateReport()
        {
            var report = new ReportHelper();
            report.CreateDoc(SetDate, ItemsCollection);
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
                // Save document
                Properties.Settings.Default.TemplatePath = dlg.FileName;
                Properties.Settings.Default.Save();
                TemplatePath = Properties.Settings.Default.TemplatePath;
                //string filename = dlg.FileName;
            }

        }
    }
}
