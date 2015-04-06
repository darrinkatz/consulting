using Microsoft.Office.Interop.Outlook;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace EmailExtractor
{
    public static class EmailExtractor
    {
        private static string _delimiter = "\t";

        public static void ProcessTable(NameSpace nameSpace, Table table, string filePath, string username)
        {
            MailItem mailItem;

            while (!table.EndOfTable)
            {
                var nextRow = table.GetNextRow();
                var item = nameSpace.GetItemFromID(nextRow["EntryID"]);

                mailItem = item as MailItem;

                if (mailItem != null)
                {
                    WriteEmailToFile(filePath, username, mailItem.Parent as Folder, mailItem);

                    Marshal.ReleaseComObject(mailItem);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private static void WriteEmailToFile(string filePath, string username, Folder folder, MailItem mailItem)
        {
            var sb = new StringBuilder();

            sb.Append(username);
            sb.Append(_delimiter);

            sb.Append(folder.FolderPath);
            sb.Append(_delimiter);

            foreach (Attachment attachment in mailItem.Attachments)
            {
                if (attachment.Type == OlAttachmentType.olEmbeddeditem)
                {
                    sb.Append(attachment.FileName);
                    sb.Append(",");
                }
            }
            sb.Append(_delimiter);

            sb.Append(mailItem.BCC);
            sb.Append(_delimiter);

            if (!string.IsNullOrEmpty(mailItem.Body))
            {
                sb.Append(new Regex(@"\s+").Replace(mailItem.Body, " "));
            }
            sb.Append(_delimiter);

            sb.Append(mailItem.CC);
            sb.Append(_delimiter);

            sb.Append(mailItem.ConversationIndex);
            sb.Append(_delimiter);

            sb.Append(mailItem.ConversationTopic);
            sb.Append(_delimiter);

            sb.Append(mailItem.CreationTime);
            sb.Append(_delimiter);

            sb.Append(mailItem.LastModificationTime);
            sb.Append(_delimiter);

            sb.Append(mailItem.ReceivedByName);
            sb.Append(_delimiter);

            sb.Append(mailItem.ReceivedTime);
            sb.Append(_delimiter);

            foreach (Recipient recipient in mailItem.Recipients)
            {
                sb.Append(string.Format("{0}|{1};", recipient.Name, recipient.Address));
            }
            sb.Append(_delimiter);

            sb.Append(mailItem.SenderEmailAddress);
            sb.Append(_delimiter);

            sb.Append(mailItem.SenderName);
            sb.Append(_delimiter);

            sb.Append(mailItem.SentOn);
            sb.Append(_delimiter);

            sb.Append(mailItem.Subject);
            sb.Append(_delimiter);

            sb.Append(mailItem.To);
            sb.Append(_delimiter);

            sb.Append(mailItem.UnRead);
            sb.Append(_delimiter);

            using (var outfile = File.AppendText(filePath))
            {
                outfile.WriteLine(sb.ToString());
            }
        }
    }
}
