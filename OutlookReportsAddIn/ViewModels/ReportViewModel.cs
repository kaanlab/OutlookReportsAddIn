using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows.Input;


namespace OutlookReportsAddIn
{
    public class ReportViewModel : BaseViewModel
    {
        public string Title { get; set; } = "Журнал принятой/отправленной корреспонденции";
        public string AddInVersion { get; private set; }
        public string AddInCompany { get; private set; }
        public string AddInCopyright { get; private set; }
        public ObservableCollection<EmailModel> ItemsCollection { get; set; } 
        public DateTime SetDate { get; set; } = DateTime.Now;

        public bool HasItems { get { return ItemsCollection?.Count > 0; } }
        public ICommand FetchItemsCommand { get; set; }
        public ICommand CreateReportCommand { get; set; }

        public ReportViewModel()
        {

            //var version = "";
            //version..GetVersion();
            AddInVersion = "ver.: " + AssemblyInfoHelper.Version.ToString();
            AddInCompany = AssemblyInfoHelper.Company.ToString();
            AddInCopyright = AssemblyInfoHelper.Copyright.ToString();

            FetchItemsCommand = new RelayCommand(FetchItems);
            CreateReportCommand = new RelayCommand(CreateReport);
        }

        private void FetchItems()
        {
            var inboxMails = MailHelper.SearchInFolder(SetDate, OlDefaultFolders.olFolderInbox);
            var sentMails = MailHelper.SearchInFolder(SetDate, OlDefaultFolders.olFolderSentMail);

            var itemsList = new List<EmailModel>();
            itemsList.AddRange(inboxMails);
            itemsList.AddRange(sentMails);

            ItemsCollection = new ObservableCollection<EmailModel>(itemsList);
        }

        private void CreateReport()
        {
            var report = new ReportHelper();
            report.CreateDoc(SetDate, ItemsCollection);
        }

    }   
}
