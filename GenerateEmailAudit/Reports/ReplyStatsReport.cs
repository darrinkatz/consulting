using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenerateEmailAudit
{
    class ReplyStatsReport : IReport
    {
        public string GetReportFileName()
        {
            return "ReplyStatsReport.txt";
        }

        public string[] Create(List<MailItem> mailItems)
        {
            var result = new List<string>();
            var sb = new StringBuilder();

            sb.Append(string.Format("{0}\t", "Time Before First Reply"));
            sb.Append(string.Format("{0}\t", "% for Emails Received"));
            sb.Append(string.Format("{0}\t", "% for Emails Sent"));
            result.Add(sb.ToString());
            sb.Clear();

            result.Add(string.Empty);

            var allEmailsReceived = MailItemHelper.GetAllEmailsReceived(mailItems);
            var allEmailsSent = MailItemHelper.GetAllEmailsSent(mailItems);

            var responsesToReceived = GetFirstResponses(mailItems, allEmailsReceived, true);
            var responsesToSent = GetFirstResponses(mailItems, allEmailsSent, false);

            GenerateReplyStatsReportLine(result, sb, responsesToReceived, responsesToSent, new TimeSpan(0, 0, 0), new TimeSpan(0, 5, 0));
            GenerateReplyStatsReportLine(result, sb, responsesToReceived, responsesToSent, new TimeSpan(0, 5, 0), new TimeSpan(0, 15, 0));
            GenerateReplyStatsReportLine(result, sb, responsesToReceived, responsesToSent, new TimeSpan(0, 15, 0), new TimeSpan(1, 0, 0));
            GenerateReplyStatsReportLine(result, sb, responsesToReceived, responsesToSent, new TimeSpan(1, 0, 0), new TimeSpan(4, 0, 0));
            GenerateReplyStatsReportLine(result, sb, responsesToReceived, responsesToSent, new TimeSpan(4, 0, 0), new TimeSpan(24, 0, 0));
            GenerateReplyStatsReportLine(result, sb, responsesToReceived, responsesToSent, new TimeSpan(24, 0, 0), TimeSpan.MaxValue);

            return result.ToArray();
        }

        private void GenerateReplyStatsReportLine(List<string> result, StringBuilder sb, List<TimeSpan> responsesReceived, List<TimeSpan> responsesSent, TimeSpan minTime, TimeSpan maxTime)
        {
            sb.Append(string.Format("{0}\t", string.Format("{0} - {1}", minTime, maxTime == TimeSpan.MaxValue ? string.Empty : maxTime.ToString())));
            sb.Append(string.Format("{0}\t", GetForTimeRange(responsesReceived, minTime, maxTime) / (double)responsesReceived.Count * 100));
            sb.Append(string.Format("{0}\t", GetForTimeRange(responsesSent, minTime, maxTime) / (double)responsesSent.Count * 100));
            result.Add(sb.ToString());
            sb.Clear();
        }

        public List<TimeSpan> GetFirstResponses(List<MailItem> allMailItems, List<MailItem> subsetMailItems, bool responseFromUserOnly)
        {
            var result = new List<TimeSpan>();

            foreach (var mailItem in subsetMailItems)
            {
                var responses = allMailItems.Where(mi =>
                    mi.ConversationIndex.Contains(mailItem.ConversationIndex)
                    && mi.ConversationIndex.Length > mailItem.ConversationIndex.Length
                    && ((responseFromUserOnly && mailItem.CurrentUserAddressEntry.Equals(mi.SenderName)) || (!responseFromUserOnly && !mailItem.CurrentUserAddressEntry.Equals(mi.SenderName)))
                    );

                if (responses.Any())
                {
                    var bestResponseTime = TimeSpan.MaxValue;

                    foreach (var response in responses)
                    {
                        var responseTime = response.CreationTime - mailItem.CreationTime;

                        if (responseTime < bestResponseTime)
                        {
                            bestResponseTime = responseTime;
                        }
                    }

                    result.Add(bestResponseTime);
                }
            }

            return result;
        }

        public int GetForTimeRange(List<TimeSpan> times, TimeSpan minTime, TimeSpan maxTime)
        {
            return times.Where(r => r >= minTime && r <= maxTime).Count();
        }
    }
}
