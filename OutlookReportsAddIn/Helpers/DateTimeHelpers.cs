using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookReportsAddIn.Helpers
{
    public static class DateTimeHelpers
    {
        public static DateTime EndOfTheMonth(this DateTime date)
        {
            var endOfTheMonth = new DateTime(date.Year, date.Month, 1)
                .AddMonths(1)
                .AddDays(-1);

            return endOfTheMonth;
        }

        public static DateTime BeginningOfTheMonth(this DateTime date)
        {
            return new DateTime(date.Year, date.Month, 1);
        }

        public static DateTime BeginningOfTheDay(this DateTime date)
        {
            return date.Date;
        }

        public static DateTime EndOfTheDay(this DateTime date)
        {
            return date.Date.AddHours(24);
        }
    }
}
