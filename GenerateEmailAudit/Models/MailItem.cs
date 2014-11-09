using System;
using System.Collections.Generic;
using System.Linq;

namespace GenerateEmailAudit
{
    public class MailItem
    {
        public string CurrentUserAddressEntry { get; set; }
        public string FolderPath { get; set; }
        public string[] Attachments { get; set; }
        public string[] BCC { get; set; }
        public string Body { get; set; }
        public string[] CC { get; set; }
        public string ConversationIndex { get; set; }
        public string ConversationTopic { get; set; }
        public DateTime CreationTime { get; set; }
        public DateTime LastModificationTime { get; set; }
        public string ReceivedByName { get; set; }
        public DateTime ReceivedTime { get; set; }
        public List<Tuple<string, string>> Recipients { get; set; }
        public string SenderEmailAddress { get; set; }
        public string SenderName { get; set; }
        public DateTime SentOn { get; set; }
        public string Subject { get; set; }
        public string[] To { get; set; }
        public bool Unread { get; set; }

        private MailItem()
        {
        }

        public static MailItem Create(string line)
        {
            MailItem result = null;

            try
            {
                var columns = line.Split('\t');

                result = new MailItem()
                {
                    CurrentUserAddressEntry = columns[0],
                    FolderPath = columns[1],
                    Attachments = columns[2].Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries),
                    BCC = columns[3].Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries).Select<string, string>(bcc => bcc.Trim()).ToArray(),
                    Body = columns[4],
                    CC = columns[5].Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries).Select<string, string>(bcc => bcc.Trim()).ToArray(),
                    ConversationIndex = columns[6],
                    ConversationTopic = columns[7],
                    CreationTime = DateTime.Parse(columns[8]),
                    LastModificationTime = DateTime.Parse(columns[9]),
                    ReceivedByName = columns[10],
                    ReceivedTime = DateTime.Parse(columns[11]),
                    Recipients = new List<Tuple<string, string>>(
                        columns[12].Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries).Select(t => new Tuple<string, string>(t.Split('|')[0], t.Split('|')[1]))
                        ),
                    SenderEmailAddress = columns[13],
                    SenderName = columns[14],
                    SentOn = DateTime.Parse(columns[15]),
                    Subject = columns[16],
                    To = columns[17].Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries).Select<string, string>(bcc => bcc.Trim()).ToArray(),
                    Unread = bool.Parse(columns[18]),
                };
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("Error Ingesting {0}: {1}", line, ex.Message));
                Console.WriteLine();
                // swallow exception and continue
            }

            return result;
        }
    }
}
