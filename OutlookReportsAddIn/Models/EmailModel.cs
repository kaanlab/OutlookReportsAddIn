using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookReportsAddIn
{
    public class EmailModel
    {
        public string SenderAddress { get; set; }

        public string Attachments { get; set; }

        public string Category { get; set; }

        public DateTime Date { get; set; }

        public string RecivedAddress { get; set; }

        public string Subject { get; set; }

    }
}
