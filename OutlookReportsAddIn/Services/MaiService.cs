using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using OutlookReportsAddIn.Extentions;

namespace OutlookReportsAddIn.Services
{
    public class MailService
    {
        public IEnumerable<Mail> SearchMails((DateTime startDate, DateTime endDate) datesRange)
        {
            var emailItems = new List<Mail>();
            emailItems.AddRange(SearchIn(datesRange, OlDefaultFolders.olFolderInbox));
            emailItems.AddRange(SearchIn(datesRange, OlDefaultFolders.olFolderSentMail));
            emailItems.Sort((x, y) => DateTime.Compare(x.Date, y.Date));

            return emailItems;
        }

        private IEnumerable<Mail> SearchIn((DateTime startDate, DateTime endDate) datesRange, OlDefaultFolders olFolder)
        {
            var emailItems = new List<Mail>();

            var stores = Globals.ThisAddIn.Application.ActiveExplorer().Session.Stores;

            foreach (Store store in stores)
            {
                var folder = store.GetDefaultFolder(olFolder);
                if (folder.FolderPath.Contains(Properties.Settings.Default.MailAddress))
                {
                    var filtredItiems = folder.Items.FiltredByDates(datesRange);

                    foreach (var item in filtredItiems)
                    {
                        if (item is MailItem)
                        {
                            var olMail = ((MailItem)item);
                            emailItems.Add(olMail.MapToMail());
                        }
                    }
                }
            }
            return emailItems;
        }
    }
}
