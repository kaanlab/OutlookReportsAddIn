using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;


namespace OutlookReportsAddIn.Services
{
    public class MailService
    {
        public IEnumerable<Mail> InboxMails((DateTime startDate, DateTime endDate) datesRange)
        {
            return SearchInFolder(datesRange, OlDefaultFolders.olFolderInbox);

        }
        public IEnumerable<Mail> OutboxMails((DateTime startDate, DateTime endDate) datesRange)
        {
            return SearchInFolder(datesRange, OlDefaultFolders.olFolderOutbox);

        }

        private IEnumerable<Mail> SearchInFolder((DateTime startDate, DateTime endDate) datesRange, OlDefaultFolders olFolder)
        {
            var emailItems = new List<Mail>();

            var accounts = Globals.ThisAddIn.Application.Session.Accounts;


            foreach (Account acc in accounts)
            {
                if (acc.AccountType == OlAccountType.olPop3)
                {
                    var stores = Globals.ThisAddIn.Application.Session.Stores;
                    foreach (Store store in stores)
                    {

                        var inboxFolder = store.GetDefaultFolder(olFolder);
                        var filtredItiems = inboxFolder.Items.
                            Restrict("[ReceivedTime] >= '" + datesRange.startDate.ToString("g") + "' And [ReceivedTime] <= '" + datesRange.endDate.ToString("g") + "'");

                        foreach (var item in filtredItiems)
                        {
                            if (item is MailItem)
                            {
                                var emailItem = new Mail();

                                // sender
                                emailItem.SenderAddress = ((MailItem)item).SenderEmailAddress;

                                // attachments
                                if (((MailItem)item).Attachments.Count > 0)
                                {
                                    var sb = new StringBuilder();

                                    for (int i = 1; i <= ((MailItem)item).Attachments.Count; i++)
                                    {
                                        sb.Append(((MailItem)item).Attachments[i].FileName + " (" + ((MailItem)item).Attachments[i].Size / 1000 + " КБ); \n");
                                        emailItem.Attachments = sb.ToString();
                                    }
                                }
                                else
                                {
                                    emailItem.Attachments = " - ";
                                }

                                //category
                                switch (((MailItem)item).Importance)
                                {
                                    case OlImportance.olImportanceNormal:
                                        emailItem.Category = "Обычная";
                                        break;
                                    case OlImportance.olImportanceHigh:
                                        emailItem.Category = "Срочно";
                                        break;
                                    default:
                                        emailItem.Category = "Низкая";
                                        break;
                                }

                                // date
                                emailItem.Date = ((MailItem)item).ReceivedTime;

                                // to
                                emailItem.RecivedAddress = ((MailItem)item).To;

                                // description
                                emailItem.Subject = "Тема сообщения: " + ((MailItem)item).Subject;

                                emailItems.Add(emailItem);
                            }
                        }
                    }
                }
            }

            return emailItems;

        }
    }
}
