using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Text;


namespace OutlookReportsAddIn
{
    public static class MailHelper
    {
        public static IEnumerable<EmailModel> SearchInFolder(DateTime date, OlDefaultFolders olFolder)
        {
            var emailItems = new List<EmailModel>();
            
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
                            Restrict("[ReceivedTime] >= '" + date.ToString("g") + "' And [ReceivedTime] < '" + date.AddDays(1).ToString("g") + "'");

                        foreach (MailItem item in filtredItiems)
                        {
                            var emailItem = new EmailModel();

                            // sender
                            emailItem.SenderAddress = item.SenderEmailAddress;

                            // attachments
                            if (item.Attachments.Count > 0)
                            {
                                var sb = new StringBuilder();

                                for (int i = 1; i <= item.Attachments.Count; i++)
                                {
                                    sb.Append(item.Attachments[i].FileName + " (" + item.Attachments[i].Size / 1000 + " КБ); \n");
                                    emailItem.Attachments = sb.ToString();
                                }
                            }
                            else
                            {
                                emailItem.Attachments = " - ";
                            }

                            //category
                            switch (item.Importance)
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
                            emailItem.Date = item.ReceivedTime;

                            // to
                            emailItem.RecivedAddress = item.To;

                            // description
                            emailItem.Subject = "Тема сообщения: " + item.Subject;
                            
                            emailItems.Add(emailItem);
                        }
                    }
                }
            }

            return emailItems;

        }
    }
}
