﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenerateEmailAudit
{
    class RecipientsReport : IReport
    {
        public string GetReportFileName()
        {
            return "RecipientsReport.txt";
        }

        public string[] Create(List<MailItem> mailItems)
        {
            var result = new List<string>();
            var sb = new StringBuilder();

            sb.Append(string.Format("{0}\t", "Recipient Name"));
            sb.Append(string.Format("{0}\t", "Recipient Email"));
            sb.Append(string.Format("{0}\t", "Total Sent"));
            sb.Append(string.Format("{0}\t", "Total To"));
            sb.Append(string.Format("{0}\t", "Total CC"));
            sb.Append(string.Format("{0}\t", "Total BCC"));
            result.Add(sb.ToString());
            sb.Clear();

            result.Add(string.Empty);

            var allEmailsSent = MailItemHelper.GetAllEmailsSent(mailItems);
            sb.Append(string.Format("{0}\t", "All Recipients"));
            sb.Append(string.Format("{0}\t", "(anyone)"));
            sb.Append(string.Format("{0}\t", allEmailsSent.Count));
            sb.Append(string.Format("{0}\t", allEmailsSent.Where(mi => mi.To.Length > 0).Count()));
            sb.Append(string.Format("{0}\t", allEmailsSent.Where(mi => mi.CC.Length > 0).Count()));
            sb.Append(string.Format("{0}\t", allEmailsSent.Where(mi => mi.BCC.Length > 0).Count()));
            result.Add(sb.ToString());
            sb.Clear();

            var allInternalEmailsSent = allEmailsSent.Where(mi => mi.Recipients.Any(r => r.Item2.Contains("@bmo.com"))).ToList();
            sb.Append(string.Format("{0}\t", "Internal Recipients"));
            sb.Append(string.Format("{0}\t", "(anyone @bmo.com)"));
            sb.Append(string.Format("{0}\t", allInternalEmailsSent.Count));
            sb.Append(string.Format("{0}\t", allInternalEmailsSent.Where(mi => mi.To.Length > 0).Count()));
            sb.Append(string.Format("{0}\t", allInternalEmailsSent.Where(mi => mi.CC.Length > 0).Count()));
            sb.Append(string.Format("{0}\t", allInternalEmailsSent.Where(mi => mi.BCC.Length > 0).Count()));
            result.Add(sb.ToString());
            sb.Clear();

            var allExternalEmailsSent = allEmailsSent.Where(mi => !mi.Recipients.Any(r => r.Item2.Contains("@bmo.com"))).ToList();
            sb.Append(string.Format("{0}\t", "External Recipients"));
            sb.Append(string.Format("{0}\t", "(anyone not @bmo.com)"));
            sb.Append(string.Format("{0}\t", allExternalEmailsSent.Count));
            sb.Append(string.Format("{0}\t", allExternalEmailsSent.Where(mi => mi.To.Length > 0).Count()));
            sb.Append(string.Format("{0}\t", allExternalEmailsSent.Where(mi => mi.CC.Length > 0).Count()));
            sb.Append(string.Format("{0}\t", allExternalEmailsSent.Where(mi => mi.BCC.Length > 0).Count()));
            result.Add(sb.ToString());
            sb.Clear();

            result.Add(string.Empty);

            var recipients = this.GetRecipients(allEmailsSent);
            foreach (var recipient in recipients)
            {
                result.Add(this.GenerateRecipientReportLine(
                    recipient.Item1,
                    recipient.Item2,
                    allEmailsSent.Where(mi =>
                        mi.To.Contains(recipient.Item1)
                        || mi.CC.Contains(recipient.Item1)
                        || mi.BCC.Contains(recipient.Item1)
                    ).ToList()
                    ));
            }

            return result.ToArray();
        }

        private List<Tuple<string, string>> GetRecipients(List<MailItem> mailItems)
        {
            var result = new List<Tuple<string, string>>();

            foreach (var mailItem in mailItems)
            {
                foreach (var recipient in mailItem.Recipients)
                {
                    if (!result.Contains(recipient))
                    {
                        result.Add(recipient);
                    }
                }
            }

            return result;
        }

        private string GenerateRecipientReportLine(string name, string email, List<MailItem> mailItems)
        {
            var sb = new StringBuilder();

            sb.Append(string.Format("{0}\t", name));
            sb.Append(string.Format("{0}\t", email));
            sb.Append(string.Format("{0}\t", mailItems.Count));
            sb.Append(string.Format("{0}\t", mailItems.Where(mi => mi.To.Contains(name)).Count()));
            sb.Append(string.Format("{0}\t", mailItems.Where(mi => mi.CC.Contains(name)).Count()));
            sb.Append(string.Format("{0}\t", mailItems.Where(mi => mi.BCC.Contains(name)).Count()));

            return sb.ToString();
        }
    }
}
