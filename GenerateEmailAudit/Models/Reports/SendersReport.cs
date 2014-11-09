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

            result.Add(string.Empty);

            var allEmailsReceived = MailItemHelper.GetAllEmailsReceived(mailItems);
            result.Add(this.GenerateSenderReportLine("All Senders", "(anyone)", allEmailsReceived));

            var allInternalEmailsReceived = allEmailsReceived.Where(mi => mi.SenderEmailAddress.Contains("@bmo.com")).ToList();
            result.Add(this.GenerateSenderReportLine("Internal Senders", "(anyone @bmo.com)", allInternalEmailsReceived));

            var allExternalEmailsReceived = allEmailsReceived.Where(mi => !mi.SenderEmailAddress.Contains("@bmo.com")).ToList();
            result.Add(this.GenerateSenderReportLine("External Senders", "(anyone not @bmo.com)", allExternalEmailsReceived));

            result.Add(string.Empty);

            var senders = this.GetSenders(allEmailsReceived);
            foreach (var sender in senders)
            {
                result.Add(this.GenerateSenderReportLine(
                    sender.Item1,
                    sender.Item2,
                    allEmailsReceived.Where(mi => mi.SenderName.Equals(sender.Item1) && mi.SenderEmailAddress.Equals(sender.Item2)).ToList())
                    );
            }

            return result.ToArray();
        }

        private List<Tuple<string, string>> GetSenders(List<MailItem> mailItems)
        {
            var result = new List<Tuple<string, string>>();

            foreach (var mailItem in mailItems)
            {
                var sender = new Tuple<string, string>(mailItem.SenderName, mailItem.SenderEmailAddress);

                if (!result.Contains(sender))
                {
                    result.Add(sender);
                }
            }

            return result;
        }


        private string GenerateSenderReportLine(string name, string email, List<MailItem> mailItems)
        {
            var sb = new StringBuilder();

            sb.Append(string.Format("{0}\t", name));
            sb.Append(string.Format("{0}\t", email));
            sb.Append(string.Format("{0}\t", mailItems.Count));
            sb.Append(string.Format("{0}\t", mailItems.Where(mi => mi.To.Contains(mi.CurrentUserAddressEntry)).Count()));
            sb.Append(string.Format("{0}\t", mailItems.Where(mi => mi.CC.Contains(mi.CurrentUserAddressEntry)).Count()));
            sb.Append(string.Format("{0}\t", mailItems.Where(mi => mi.BCC.Contains(mi.CurrentUserAddressEntry)).Count()));

            return sb.ToString();
        }
    }
}
