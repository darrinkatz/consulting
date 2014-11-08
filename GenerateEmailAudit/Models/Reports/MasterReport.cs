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

            foreach (var mailItem in mailItems)
            {
                var sb = new StringBuilder();

                sb.Append(mailItem.CurrentUserAddressEntry);
                sb.Append("\t");
                sb.Append(mailItem.FolderPath);
                sb.Append("\t");
                sb.Append(string.Join(",", mailItem.Attachments));
                sb.Append("\t");
                sb.Append(mailItem.BCC);
                sb.Append("\t");
                sb.Append(mailItem.Body);
                sb.Append("\t");
                sb.Append(mailItem.CC);
                sb.Append("\t");
                sb.Append(mailItem.ConversationIndex);
                sb.Append("\t");
                sb.Append(mailItem.ConversationTopic);
                sb.Append("\t");
                sb.Append(mailItem.CreationTime);
                sb.Append("\t");
                sb.Append(mailItem.LastModificationTime);
                sb.Append("\t");
                sb.Append(mailItem.ReceivedByName);
                sb.Append("\t");
                sb.Append(mailItem.ReceivedTime);
                sb.Append("\t");
                sb.Append(string.Join(";", mailItem.Recipients));
                sb.Append("\t");
                sb.Append(mailItem.SenderEmailAddress);
                sb.Append("\t");
                sb.Append(mailItem.SenderName);
                sb.Append("\t");
                sb.Append(mailItem.SentOn);
                sb.Append("\t");
                sb.Append(mailItem.Subject);
                sb.Append("\t");
                sb.Append(string.Join(";", mailItem.To));
                sb.Append("\t");
                sb.Append(mailItem.Unread);
                sb.Append("\t");

                result.Add(sb.ToString());
            }

            return result.ToArray();
        }
    }
}
