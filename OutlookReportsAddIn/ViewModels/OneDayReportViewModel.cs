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

        public DateTime SelectedDate { get; set; } = DateTime.Now;

        public ICommand FetchItemsCommand { get; }
        public ICommand CreateReportCommand { get; }
        public ICommand SetTemplatePathCommand { get; }


        public OneDayReportViewModel()
        {
            _mailService = new MailService();
            _exportService = new ExportService();

            FetchItemsCommand = new RelayCommand(FetchItems);
            CreateReportCommand = new RelayCommand(CreateReport);
            SetTemplatePathCommand = new RelayCommand(SetTemplate);
        }

        private void FetchItems()
        {
            var datesRange = (SelectedDate.BeginningOfTheDay(), SelectedDate.EndOfTheDay());
            var inboxMails = _mailService.InboxMails(datesRange);
            var sentMails = _mailService.OutboxMails(datesRange);

            var itemsList = new List<Mail>();
            itemsList.AddRange(inboxMails);
            itemsList.AddRange(sentMails);
            itemsList.Sort((x, y) => DateTime.Compare(x.Date, y.Date));

            ItemsCollection = new ObservableCollection<Mail>(itemsList);
            HasItems = ItemsCollection?.Count > 0;
        }

        private void CreateReport()
        {
            _exportService.ToWord(SelectedDate, ItemsCollection, MailsCounter);
            MailsCounter += ItemsCollection.Count;
            Properties.Settings.Default.MailCounter = MailsCounter;
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
                // Save document
                Properties.Settings.Default.TemplatePath = dlg.FileName;
                Properties.Settings.Default.Save();
                TemplatePath = Properties.Settings.Default.TemplatePath;
                IsTemplatePathExsist = true; // update image     
            }

        }
    }
}
