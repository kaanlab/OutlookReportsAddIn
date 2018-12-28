using System;

namespace OutlookReportsAddIn
{
    public class Mail
    {
        public string SenderAddress { get; set; }

        public string Attachments { get; set; }

        public string Category { get; set; }

        public DateTime Date { get; set; }

        public string RecivedAddress { get; set; }

        public string Subject { get; set; }

    }
}
