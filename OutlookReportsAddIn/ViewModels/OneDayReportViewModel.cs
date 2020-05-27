using OutlookReportsAddIn.Helpers;
using OutlookReportsAddIn.Services;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows.Input;

namespace OutlookReportsAddIn.ViewModels
{
    public class OneDayReportViewModel : BaseViewModel
    {
        private readonly MailService _mailService;

        private readonly ExportService _exportService;

        public static string WindowTitle { get => "Журнал принятой/отправленной корреспонденции за день"; }

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

        public DateTime SelectedDate { get; set; } = DateTime.Now;
        public ICommand FetchItemsCommand { get; }
        public ICommand CreateReportCommand { get; }

        public OneDayReportViewModel()
        {
            _mailService = new MailService();
            _exportService = new ExportService();

            FetchItemsCommand = new RelayCommand(FetchItems);
            CreateReportCommand = new RelayCommand(CreateReport);
        }

        private void FetchItems()
        {
            var datesRange = (SelectedDate.BeginningOfTheDay(), SelectedDate.EndOfTheDay());
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
