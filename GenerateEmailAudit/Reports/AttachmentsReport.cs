using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenerateEmailAudit
{
    class AttachmentsReport : IReport
    {
        public string GetReportFileName()
        {
            return "AttachmentsReport.txt";
        }

        public string[] Create(List<MailItem> mailItems)
        {
            var result = new List<string>();
            var sb = new StringBuilder();

            sb.Append(string.Format("{0}\t", "Emails with Attachments"));
            sb.Append(string.Format("{0}\t", ".pdf"));
            sb.Append(string.Format("{0}\t", ".xls"));
            sb.Append(string.Format("{0}\t", ".doc"));
            sb.Append(string.Format("{0}\t", "other"));
            result.Add(sb.ToString());
            sb.Clear();

            result.Add(string.Empty);

            var allEmailsSent = MailItemHelper.GetAllEmailsSent(mailItems);
            var allEmailsReceived = MailItemHelper.GetAllEmailsReceived(mailItems);

            var allInternalEmailsSent = allEmailsSent.Where(mi => mi.Recipients.Any(r => r.Item2.Contains("@bmo.com"))).ToList();
            result.Add(GenerateRecipientReportLine("Sent to anyone @bmo.com", allInternalEmailsSent));
            sb.Clear();

            var allExternalEmailsSent = allEmailsSent.Where(mi => !mi.Recipients.Any(r => r.Item2.Contains("@bmo.com"))).ToList();
            result.Add(GenerateRecipientReportLine("Sent to anyone not @bmo.com", allExternalEmailsSent));
            sb.Clear();

            result.Add(string.Empty);

            var allInternalEmailsReceived = allEmailsReceived.Where(mi => mi.SenderEmailAddress.Contains("@bmo.com")).ToList();
            result.Add(GenerateRecipientReportLine("Received by anyone @bmo.com", allInternalEmailsReceived));
            sb.Clear();

            var allExternalEmailsReceived = allEmailsReceived.Where(mi => !mi.SenderEmailAddress.Contains("@bmo.com")).ToList();
            result.Add(GenerateRecipientReportLine("Received by anyone not @bmo.com", allExternalEmailsReceived));
            sb.Clear();

            return result.ToArray();
        }

        private string GenerateRecipientReportLine(string label, List<MailItem> mailItems)
        {
            var sb = new StringBuilder();

            sb.Append(string.Format("{0}\t", label));
            sb.Append(string.Format("{0}\t", mailItems.Where(mi => mi.Attachments.Any(a => a.Contains(".pdf"))).Count()));
            sb.Append(string.Format("{0}\t", mailItems.Where(mi => mi.Attachments.Any(a => a.Contains(".xls"))).Count()));
            sb.Append(string.Format("{0}\t", mailItems.Where(mi => mi.Attachments.Any(a => a.Contains(".doc"))).Count()));
            sb.Append(string.Format("{0}\t", mailItems.Where(mi => mi.Attachments.Any(a => !a.Contains(".pdf") && !a.Contains(".xls") && !a.Contains(".doc"))).Count()));
            return sb.ToString();
        }
    }
}
