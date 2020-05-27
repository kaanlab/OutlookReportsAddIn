using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookReportsAddIn.Extentions
{
    public static class ModelMappingExtentions
    {
        public static Mail MapToMail(this MailItem outlookMail)
        {
            var emailItem = new Mail();

            // sender
            emailItem.SenderAddress = outlookMail.SenderEmailAddress;

            // attachments
            if (outlookMail.Attachments.Count > 0)
            {
                var sb = new StringBuilder();

                for (int i = 1; i <= outlookMail.Attachments.Count; i++)
                {
                    sb.Append(outlookMail.Attachments[i].FileName + " (" + outlookMail.Attachments[i].Size / 1000 + " КБ); \n");
                    emailItem.Attachments = sb.ToString();
                }
            }
            else
            {
                emailItem.Attachments = " - ";
            }

            //category
            switch (outlookMail.Importance)
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
            emailItem.Date = outlookMail.ReceivedTime;

            // to
            emailItem.RecivedAddress = outlookMail.To;

            // description
            emailItem.Subject = "Тема сообщения: " + outlookMail.Subject;

            return emailItem;
        }
    }
}
