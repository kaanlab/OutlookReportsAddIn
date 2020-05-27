using OutlookReportsAddIn.Helpers;
using OutlookReportsAddIn.Services;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace OutlookReportsAddIn.ViewModels
{
    public class ReportForPeriodViewModel : BaseViewModel
    {
        private readonly MailService _mailService;

        private readonly ExportService _exportService;

        public static string WindowTitle { get => "Журнал принятой/отправленной корреспонденции за период"; }

        private ObservableCollection<Mail> _itemsCollection;
        public ObservableCollection<Mail> ItemsCollection
        {
            get => _itemsCollection;
            set
            {
                _itemsCollection = value;
                OnPropertyChanged("ItemsCollection");
            }
        }

        private bool _hasItems;
        public bool HasItems
        {
            get => _hasItems;
            set
            {
                _hasItems = value;
                OnPropertyChanged("HasItems");
            }
        }

        private int _mailsCounter = Properties.Settings.Default.MailCounter;
        public int MailsCounter
        {
            get => _mailsCounter;
            set
            {
                _mailsCounter = value;
                OnPropertyChanged("MailsCounter");
            }
        }

        public DateTime StartDate { get; set; } = DateTime.Now.BeginningOfTheMonth();
        public DateTime EndDate { get; set; } = DateTime.Now.EndOfTheMonth();
        public ICommand FetchItemsCommand { get; }
        public ICommand CreateReportCommand { get; }

        public ReportForPeriodViewModel()
        {
            _mailService = new MailService();
            _exportService = new ExportService();

            FetchItemsCommand = new RelayCommand(FetchItems);
            CreateReportCommand = new RelayCommand(CreateReport);
        }

        private void FetchItems()
        {
            var datesRange = (StartDate.BeginningOfTheDay(), EndDate.EndOfTheDay());
            ItemsCollection = new ObservableCollection<Mail>(_mailService.SearchMails(datesRange));
            HasItems = ItemsCollection?.Count > 0;
        }

        private void CreateReport()
        {
            _exportService.ToWord(ItemsCollection, MailsCounter);
            MailsCounter += ItemsCollection.Count;
            Properties.Settings.Default.MailCounter = MailsCounter;
            Properties.Settings.Default.Save();
        }
    }
}
