using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookReportsAddIn.Extentions
{
    public static class FiltredItemsExtentions
    {
        public static Items FiltredByDates(this Items items, (DateTime startDate, DateTime endDate) datesRange)
        {
            return items.Restrict("[ReceivedTime] >= '" + datesRange.startDate.ToString("g") + "' And [ReceivedTime] <= '" + datesRange.endDate.ToString("g") + "'");
        }
    }
}
