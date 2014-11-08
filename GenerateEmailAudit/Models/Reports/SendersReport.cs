using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenerateEmailAudit
{
    class SendersReport : IReport
    {
        public string GetReportFileName()
        {
            return "SendersReport.txt";
        }

        public string[] Create(List<MailItem> mailItems)
        {
            var result = new List<string>();
            var sb = new StringBuilder();

            sb.Append(string.Format("{0}\t", "Sender Name"));
            sb.Append(string.Format("{0}\t", "Sender Email"));
            sb.Append(string.Format("{0}\t", "Total Received"));
            sb.Append(string.Format("{0}\t", "Total To"));
            sb.Append(string.Format("{0}\t", "Total CC"));
            sb.Append(string.Format("{0}\t", "Total BCC"));
            result.Add(sb.ToString());
            sb.Clear();

            result.Add(sb.ToString());

            var allEmailsReceived = MailItem.GetAllEmailsReceived(mailItems);
            sb.Append(string.Format("{0}\t", "All Senders"));
            sb.Append(string.Format("{0}\t", "(anyone)"));
            sb.Append(string.Format("{0}\t", allEmailsReceived.Count));
            sb.Append(string.Format("{0}\t", "Total To"));
            sb.Append(string.Format("{0}\t", "Total CC"));
            sb.Append(string.Format("{0}\t", "Total BCC"));
            result.Add(sb.ToString());
            sb.Clear();

            return result.ToArray();
        }
    }
}
