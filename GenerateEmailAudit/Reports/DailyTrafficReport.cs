using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GenerateEmailAudit
{
    class DailyTrafficReport : IReport
    {
        public string GetReportFileName()
        {
            return "DailyTrafficReport.txt";
        }

        public string[] Create(List<MailItem> mailItems)
        {
            var result = new List<string>();
            var sb = new StringBuilder();

            sb.Append(string.Format("{0}\t", "Time Period"));
            sb.Append(string.Format("{0}\t", "Average Emails Sent"));
            sb.Append(string.Format("{0}\t", "Average Emails Received"));
            result.Add(sb.ToString());
            sb.Clear();

            result.Add(string.Empty);

            var allEmailsSent = MailItemHelper.GetAllEmailsSent(mailItems);
            var allEmailsReceived = MailItemHelper.GetAllEmailsReceived(mailItems);
            
            var numberOfDays = GetDayCount(mailItems);
            sb.Append(string.Format("{0}\t", "Daily"));
            sb.Append(string.Format("{0}\t", allEmailsSent.Count() / numberOfDays));
            sb.Append(string.Format("{0}\t", allEmailsReceived.Count() / numberOfDays));
            result.Add(sb.ToString());
            sb.Clear();

            var numberOfWeeks = GetWeekCount(mailItems);
            sb.Append(string.Format("{0}\t", "Weekly"));
            sb.Append(string.Format("{0}\t", allEmailsSent.Count() / numberOfWeeks));
            sb.Append(string.Format("{0}\t", allEmailsReceived.Count() / numberOfWeeks));
            result.Add(sb.ToString());
            sb.Clear();

            result.Add(string.Empty);

            GenerateDailyTrafficReportLine(mailItems, result, sb, DayOfWeek.Monday);
            GenerateDailyTrafficReportLine(mailItems, result, sb, DayOfWeek.Tuesday);
            GenerateDailyTrafficReportLine(mailItems, result, sb, DayOfWeek.Wednesday);
            GenerateDailyTrafficReportLine(mailItems, result, sb, DayOfWeek.Thursday);
            GenerateDailyTrafficReportLine(mailItems, result, sb, DayOfWeek.Friday);
            GenerateDailyTrafficReportLine(mailItems, result, sb, DayOfWeek.Saturday);
            GenerateDailyTrafficReportLine(mailItems, result, sb, DayOfWeek.Sunday);

            result.Add(string.Empty);

            for (int hourOfDay = 0; hourOfDay < 24; hourOfDay++)
            {
                var allEmailsSentDuringHourOfDay = GetAllEmailsSent(mailItems, hourOfDay);
                var allEmailsReceivedDuringHourOfDay = GetAllEmailsReceived(mailItems, hourOfDay);
                sb.Append(string.Format("{0}:00 - {1}:00\t", hourOfDay, hourOfDay + 1));
                sb.Append(string.Format("{0}\t", allEmailsSentDuringHourOfDay.Count() / numberOfDays));
                sb.Append(string.Format("{0}\t", allEmailsReceivedDuringHourOfDay.Count() / numberOfDays));
                result.Add(sb.ToString());
                sb.Clear();
            }

            return result.ToArray();
        }

        private static void GenerateDailyTrafficReportLine(List<MailItem> mailItems, List<string> result, StringBuilder sb, DayOfWeek dayOfWeek)
        {
            var allEmailsSentOnWeekday = GetAllEmailsSent(mailItems, dayOfWeek);
            var allEmailsReceivedOnWeekday = GetAllEmailsReceived(mailItems, dayOfWeek);
            var numberOfWeekdays = GetDayOfWeekCount(mailItems, dayOfWeek);
            sb.Append(string.Format("{0}s\t", dayOfWeek.ToString()));
            sb.Append(string.Format("{0}\t", numberOfWeekdays > 0 ? allEmailsSentOnWeekday.Count() / numberOfWeekdays : 0));
            sb.Append(string.Format("{0}\t", numberOfWeekdays > 0 ? allEmailsReceivedOnWeekday.Count() / numberOfWeekdays : 0));
            result.Add(sb.ToString());
            sb.Clear();
        }

        private static List<MailItem> GetAllEmailsSent(List<MailItem> mailItems, DayOfWeek dayOfWeek)
        {
            return MailItemHelper.GetAllEmailsSent(mailItems.Where(mi =>
                mi.CreationTime.DayOfWeek == dayOfWeek
                ).ToList());
        }

        private static List<MailItem> GetAllEmailsSent(List<MailItem> mailItems, int hourOfDay)
        {
            return MailItemHelper.GetAllEmailsSent(mailItems.Where(mi =>
                mi.CreationTime.Hour == hourOfDay
                ).ToList());
        }

        private static List<MailItem> GetAllEmailsReceived(List<MailItem> mailItems, DayOfWeek dayOfWeek)
        {
            return MailItemHelper.GetAllEmailsReceived(mailItems.Where(mi =>
                mi.CreationTime.DayOfWeek == dayOfWeek
                ).ToList());
        }

        private static List<MailItem> GetAllEmailsReceived(List<MailItem> mailItems, int hourOfDay)
        {
            return MailItemHelper.GetAllEmailsReceived(mailItems.Where(mi =>
                mi.CreationTime.Hour == hourOfDay
                ).ToList());
        }

        private static double GetDayCount(List<MailItem> mailItems)
        {
            var dayList = new List<DateTime>();

            foreach (var mailItem in mailItems)
            {
                var day = new DateTime(
                    mailItem.CreationTime.Year,
                    mailItem.CreationTime.Month,
                    mailItem.CreationTime.Day
                    );

                if (!dayList.Contains(day))
                {
                    dayList.Add(day);
                }
            }

            return dayList.Count();
        }

        private static double GetDayOfWeekCount(List<MailItem> mailItems, DayOfWeek dayOfWeek)
        {
            return GetDayCount(mailItems.Where(mi =>
                mi.CreationTime.DayOfWeek == dayOfWeek
                ).ToList());
        }

        private static double GetWeekCount(List<MailItem> mailItems)
        {
            var weekList = new List<int>();

            foreach (var mailItem in mailItems)
            {
                int week = (int)Math.Floor((double)mailItem.CreationTime.DayOfYear / 7.0);

                if (!weekList.Contains(week))
                {
                    weekList.Add(week);
                }
            }

            return weekList.Count();
        }
    }
}
