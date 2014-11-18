using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GenerateEmailAudit
{
    class WordStatsReport : IReport
    {
        public string GetReportFileName()
        {
            return "WordStatsReport.txt";
        }

        public string[] Create(List<MailItem> mailItems)
        {
            var result = new List<string>();
            var sb = new StringBuilder();

            sb.Append(string.Format("{0}\t", "Word Stats"));
            sb.Append(string.Format("{0}\t", "Emails Sent"));
            sb.Append(string.Format("{0}\t", "Emails Received"));
            result.Add(sb.ToString());
            sb.Clear();

            result.Add(string.Empty);

            var allEmailsSent = MailItemHelper.GetAllEmailsSent(mailItems);
            var allEmailsReceived = MailItemHelper.GetAllEmailsReceived(mailItems);
            var wordsSent = GetWords(allEmailsSent);
            var wordsRecieved = GetWords(allEmailsReceived);

            sb.Append(string.Format("{0}\t", "Average Words per Email"));
            sb.Append(string.Format("{0}\t", wordsSent.Sum() / (double)allEmailsSent.Count()));
            sb.Append(string.Format("{0}\t", wordsRecieved.Sum() / (double)allEmailsReceived.Count()));
            result.Add(sb.ToString());
            sb.Clear();

            result.Add(string.Empty);

            GenerateWordStatsReportLine(result, sb, allEmailsSent, allEmailsReceived, wordsSent, wordsRecieved, 0, 9);
            GenerateWordStatsReportLine(result, sb, allEmailsSent, allEmailsReceived, wordsSent, wordsRecieved, 10, 29);
            GenerateWordStatsReportLine(result, sb, allEmailsSent, allEmailsReceived, wordsSent, wordsRecieved, 30, 49);
            GenerateWordStatsReportLine(result, sb, allEmailsSent, allEmailsReceived, wordsSent, wordsRecieved, 50, 99);
            GenerateWordStatsReportLine(result, sb, allEmailsSent, allEmailsReceived, wordsSent, wordsRecieved, 100, 199);
            GenerateWordStatsReportLine(result, sb, allEmailsSent, allEmailsReceived, wordsSent, wordsRecieved, 200, 499);
            GenerateWordStatsReportLine(result, sb, allEmailsSent, allEmailsReceived, wordsSent, wordsRecieved, 500, 999);
            GenerateWordStatsReportLine(result, sb, allEmailsSent, allEmailsReceived, wordsSent, wordsRecieved, 1000, int.MaxValue);

            return result.ToArray();
        }

        private static void GenerateWordStatsReportLine(List<string> result, StringBuilder sb, List<MailItem> allEmailsSent, List<MailItem> allEmailsReceived, List<int> wordsSent, List<int> wordsRecieved, int min, int max)
        {
            sb.Append(string.Format("{0}\t", string.Format("{0} - {1} words", min, max == int.MaxValue ? string.Empty : max.ToString())));
            sb.Append(string.Format("{0}\t", wordsSent.Where(w => w >= min && w <= max).Count() / (double)allEmailsSent.Count() * 100));
            sb.Append(string.Format("{0}\t", wordsRecieved.Where(w => w >= min && w <= max).Count() / (double)allEmailsReceived.Count() * 100));
            result.Add(sb.ToString());
            sb.Clear();
        }

        private List<int> GetWords(List<MailItem> mailItems)
        {
            var result = new List<int>();

            foreach (var mailItem in mailItems)
            {
                var words = mailItem.Body.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                result.Add(words.Length);
            }

            return result;
        }
    }
}
