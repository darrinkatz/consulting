using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace EmailExtractor
{
    public static class EmailExtractor
    {
        private static string _delimiter = "\t";

        public static void ProcessFolder(string pathFormat, string username, Folder folder)
        {
            MailItem mailItem;

            for (int i = 1; i < folder.Items.Count; i++)
            {
                mailItem = folder.Items[i] as MailItem;

                if (mailItem != null)
                {
                    //TODO: make date range parameters 
                    if (mailItem.CreationTime > new DateTime(2015, 04, 01) && mailItem.CreationTime < new DateTime(2015, 04, 06))
                    {
                        WriteEmailToFile(pathFormat, username, folder, mailItem);
                    }

                    Marshal.ReleaseComObject(mailItem);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            Folder childFolder;

            for (int i = 1; i < folder.Folders.Count; i++)
            {
                childFolder = (Folder)folder.Folders[i];

                //ProcessFolder(pathFormat, username, childFolder);
            }
        }

        private static void WriteEmailToFile(string pathFormat, string username, Folder folder, MailItem mailItem)
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

            using (var outfile = File.AppendText(string.Format(pathFormat, "data.csv")))
            {
                outfile.WriteLine(sb.ToString());
            }
        }
    }
}
