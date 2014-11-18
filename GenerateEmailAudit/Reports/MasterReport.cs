using System.Collections.Generic;
using System.Text;

namespace GenerateEmailAudit
{
    class MasterReport : IReport
    {
        public string GetReportFileName()
        {
            return "MasterReport.txt";
        }

        public string[] Create(List<MailItem> mailItems)
        {
            var result = new List<string>();
            var sb = new StringBuilder();

            sb.Append(string.Format("{0}\t", "CurrentUserAddressEntry"));
            sb.Append(string.Format("{0}\t", "FolderPath"));
            sb.Append(string.Format("{0}\t", "Attachments"));
            sb.Append(string.Format("{0}\t", "BCC"));
            sb.Append(string.Format("{0}\t", "Body"));
            sb.Append(string.Format("{0}\t", "CC"));
            sb.Append(string.Format("{0}\t", "ConversationIndex"));
            sb.Append(string.Format("{0}\t", "ConversationTopic"));
            sb.Append(string.Format("{0}\t", "CreationTime"));
            sb.Append(string.Format("{0}\t", "LastModificationTime"));
            sb.Append(string.Format("{0}\t", "ReceivedByName"));
            sb.Append(string.Format("{0}\t", "ReceivedTime"));
            sb.Append(string.Format("{0}\t", "Recipients"));
            sb.Append(string.Format("{0}\t", "SenderEmailAddress"));
            sb.Append(string.Format("{0}\t", "SenderName"));
            sb.Append(string.Format("{0}\t", "SentOn"));
            sb.Append(string.Format("{0}\t", "Subject"));
            sb.Append(string.Format("{0}\t", "To"));
            sb.Append(string.Format("{0}\t", "Unread"));

            result.Add(sb.ToString());
            sb.Clear();

            foreach (var mailItem in mailItems)
            {
                sb.Append(string.Format("{0}\t", mailItem.CurrentUserAddressEntry));
                sb.Append(string.Format("{0}\t", mailItem.FolderPath));
                sb.Append(string.Format("{0}\t", string.Join(",", mailItem.Attachments)));
                sb.Append(string.Format("{0}\t", string.Join(";", mailItem.BCC)));
                sb.Append(string.Format("{0}\t", mailItem.Body));
                sb.Append(string.Format("{0}\t", string.Join(";", mailItem.CC)));
                sb.Append(string.Format("{0}\t", mailItem.ConversationIndex));
                sb.Append(string.Format("{0}\t", mailItem.ConversationTopic));
                sb.Append(string.Format("{0}\t", mailItem.CreationTime));
                sb.Append(string.Format("{0}\t", mailItem.LastModificationTime));
                sb.Append(string.Format("{0}\t", mailItem.ReceivedByName));
                sb.Append(string.Format("{0}\t", mailItem.ReceivedTime));
                sb.Append(string.Format("{0}\t", string.Join(";", mailItem.Recipients)));
                sb.Append(string.Format("{0}\t", mailItem.SenderEmailAddress));
                sb.Append(string.Format("{0}\t", mailItem.SenderName));
                sb.Append(string.Format("{0}\t", mailItem.SentOn));
                sb.Append(string.Format("{0}\t", mailItem.Subject));
                sb.Append(string.Format("{0}\t", string.Join(";", mailItem.To)));
                sb.Append(string.Format("{0}\t", mailItem.Unread));

                result.Add(sb.ToString());
                sb.Clear();
            }

            return result.ToArray();
        }
    }
}
