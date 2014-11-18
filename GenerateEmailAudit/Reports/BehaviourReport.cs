using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenerateEmailAudit
{
    class BehaviourReport : IReport
    {
        public string GetReportFileName()
        {
            return "BehaviourReport.txt";
        }

        public string[] Create(List<MailItem> mailItems)
        {
            var result = new List<string>();
            var sb = new StringBuilder();

            sb.Append(string.Format("{0}\t", "Unopened Subject"));
            sb.Append(string.Format("{0}\t", "Unopened Sender"));
            result.Add(sb.ToString());
            sb.Clear();

            result.Add(string.Empty);

            foreach (var mailItem in mailItems)
            {
                if (mailItem.Unread)
                {
                    sb.Append(string.Format("{0}\t", mailItem.Subject));
                    sb.Append(string.Format("{0}\t", mailItem.SenderEmailAddress));
                    result.Add(sb.ToString());
                    sb.Clear();
                }
            }

            return result.ToArray();
        }
    }
}
