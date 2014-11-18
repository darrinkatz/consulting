using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenerateEmailAudit
{
    class ReportDocumentation : IReport
    {
        public string GetReportFileName()
        {
            return "ReportDocumentation.txt";
        }

        public string[] Create(List<MailItem> mailItems)
        {
            var result = new List<string>();
            var sb = new StringBuilder();

            sb.Append(string.Format("{0}\t", "Attachments Report"));
            sb.Append(string.Format("{0}\t", "If the same email contains more than one of the listed attachment types it will be counted in each column."));
            result.Add(sb.ToString());
            sb.Clear();

            sb.Append(string.Format("{0}\t", "Behaviour Report"));
            sb.Append(string.Format("{0}\t", "This report lists unopened emails."));
            result.Add(sb.ToString());
            sb.Clear();

            sb.Append(string.Format("{0}\t", "Conversations Report"));
            sb.Append(string.Format("{0}\t", "This report counts replies but does not count the initial email that started the conversation."));
            result.Add(sb.ToString());
            sb.Clear();

            sb.Append(string.Format("{0}\t", "Daily Traffic Report"));
            sb.Append(string.Format("{0}\t", "This report contains the average number of emails sent and received broken down by time period."));
            result.Add(sb.ToString());
            sb.Clear();

            sb.Append(string.Format("{0}\t", "Master Report"));
            sb.Append(string.Format("{0}\t", "This report contains all of the raw data captured from the user's emails."));
            result.Add(sb.ToString());
            sb.Clear();

            sb.Append(string.Format("{0}\t", "Recipients Report"));
            sb.Append(string.Format("{0}\t", "This report contains a list of the email addresses sent to and the frequency for each."));
            result.Add(sb.ToString());
            sb.Clear();

            sb.Append(string.Format("{0}\t", "Reply Stats Report"));
            sb.Append(string.Format("{0}\t", "This report only considers the first reponse to emails that have been replied to."));
            result.Add(sb.ToString());
            sb.Clear();

            sb.Append(string.Format("{0}\t", "Report Documentation"));
            sb.Append(string.Format("{0}\t", "This report describes each of the reports."));
            result.Add(sb.ToString());
            sb.Clear();

            sb.Append(string.Format("{0}\t", "Senders Report"));
            sb.Append(string.Format("{0}\t", "This report contains a list of the email addresses received from and the frequency for each."));
            result.Add(sb.ToString());
            sb.Clear();

            sb.Append(string.Format("{0}\t", "Word Stats Report"));
            sb.Append(string.Format("{0}\t", "This report counts all words in an email including the contents of earlier emails from the thread that are copied inside the body."));
            result.Add(sb.ToString());
            sb.Clear();

            return result.ToArray();
        }
    }
}
